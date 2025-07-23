// script.js

// --- DATA SIMULASI EXCEL ---
// Ini adalah data yang akan menjadi referensi untuk perhitungan expectedResult
// Kolom dimulai dari indeks 0 (A=0, B=1, dst.)
// Baris dimulai dari indeks 0 (header), jadi data dimulai dari indeks 1.
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

// Data untuk Sheet2
const sheet2Data = [
  ["Kode Kategori", "Deskripsi Kategori"],
  ["EL", "Elektronik"],
  ["AK", "Aksesoris"],
  ["KO", "Komponen"],
];

// --- FUNGSI SIMULASI EXCEL UNTUK MENGHITUNG EXPECTED RESULT ---
// Fungsi ini akan digunakan untuk menghitung nilai expectedResult secara dinamis jika perlu
// dan juga untuk "mengerti" bagaimana Excel bekerja pada data simulasi kita.

// Helper untuk mendapatkan nilai dari sel
function getVal(rowIdx, colIdx) {
  if (excelData[rowIdx] && excelData[rowIdx][colIdx] !== undefined) {
    return excelData[rowIdx][colIdx];
  }
  return undefined;
}

// Helper untuk mendapatkan nilai dari rentang (khusus untuk fungsi agregasi)
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
        // Hanya masukkan angka atau string tidak kosong
        values.push(dataSrc[r][c]);
      }
    }
  }
  return values;
}

// Implementasi fungsi Excel yang sebenarnya
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
    if (numbers.length === 0) return 0; // Atau -Infinity
    return Math.max(...numbers);
  },
  MIN: (range) => {
    const numbers = getRangeVals(
      range.split(":")[0],
      range.split(":")[1]
    ).filter((val) => typeof val === "number");
    if (numbers.length === 0) return 0; // Atau Infinity
    return Math.min(...numbers);
  },
  IF: (condition, valueIfTrue, valueIfFalse) => {
    // Untuk IF, kita akan menganggap kondisi sudah dievaluasi.
    // Dalam soal, kita akan menyediakan kondisi yang sudah "siap".
    return condition ? valueIfTrue : valueIfFalse;
  },
  VLOOKUP: (
    lookup_value,
    table_array_ref,
    col_index_num,
    range_lookup = false
  ) => {
    // table_array_ref: contoh 'Sheet2!$A$1:$B$4'
    const tableParts = table_array_ref.split("!")[1].split(":");
    const tableStartRef = tableParts[0].replace(/\$/g, ""); // Hapus $
    const tableEndRef = tableParts[1].replace(/\$/g, ""); // Hapus $

    const tableData = getRangeVals(tableStartRef, tableEndRef, sheet2Data); // Gunakan sheet2Data
    const lookupColIdx = tableStartRef.charCodeAt(0) - "A".charCodeAt(0);
    const returnColIdx =
      tableStartRef.charCodeAt(0) - "A".charCodeAt(0) + (col_index_num - 1);

    for (let i = 0; i < sheet2Data.length; i++) {
      // Iterasi semua baris di sheet2Data
      // Cari lookup_value di kolom pertama dari table_array
      if (sheet2Data[i][lookupColIdx] === lookup_value) {
        return sheet2Data[i][returnColIdx];
      }
    }
    return "#N/A"; // Jika tidak ditemukan
  },
  CONCAT: (val1, val2, val3) => {
    // Contoh untuk 3 argumen, bisa diperluas
    return `${val1}${val2}${val3}`;
  },
  MONTH: (dateString) => {
    const date = new Date(dateString);
    return date.getMonth() + 1; // getMonth() is 0-indexed
  },
  TODAY: () => {
    const today = new Date();
    // Hanya tanggal, tanpa waktu untuk konsistensi dengan Excel
    return new Date(today.getFullYear(), today.getMonth(), today.getDate());
  },
  DATEDIF: (startDateObj, endDateObj, unit) => {
    const diffTime = Math.abs(endDateObj.getTime() - startDateObj.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    if (unit.toUpperCase() === "D") return diffDays;
    return diffDays; // Untuk kesederhanaan, hanya implementasi hari
  },
};

// --- DATA SOAL PER LEVEL ---
// expectedResult sekarang dihitung secara dinamis atau hardcoded jika kompleks
const questions = [
  // Level 1: Dasar Matematika & Statistik Sederhana
  {
    level: 1, // 'level' ini bisa dihapus atau diabaikan sekarang
    question:
      "Berapa **Total Harga (Sebelum Diskon)** untuk 'Laptop XYZ'? (Rumus di sel H2: D2*E2)",
    resultCell: "H2",
    expectedResult: () => getVal(1, 3) * getVal(1, 4), // D2 * E2
    type: "number",
  },
  {
    level: 1,
    question:
      "Berapa **Total Harga (Setelah Diskon)** untuk 'Laptop XYZ'? (Rumus di sel I2: H2-(H2*F2))",
    resultCell: "I2",
    expectedResult: () =>
      getVal(1, 3) * getVal(1, 4) - getVal(1, 3) * getVal(1, 4) * getVal(1, 5), // (D2*E2) - ((D2*E2)*F2)
    type: "number",
  },
  {
    level: 1,
    question:
      "Berapa **total penjualan semua produk sebelum diskon**? (Rumus di sel H10: SUM(H2:H9))",
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
    level: 1,
    question:
      "Berapa **rata-rata 'Harga Satuan'** dari semua produk? (Rumus di sel D10: AVERAGE(D2:D9))",
    resultCell: "D10",
    expectedResult: () => calculateExcelFunction.AVERAGE("D2:D9"),
    type: "number",
  },
  // Level 2: Logika & Teks (Sekarang jadi Nomor Soal berurutan)
  {
    level: 2, // Ini juga bisa diabaikan
    question:
      "Apa **'Status Diskon'** untuk 'Laptop XYZ'? (Rumus di sel J2: IF(F2>0,\"Ada Diskon\",\"Tidak Ada Diskon\"))",
    resultCell: "J2",
    expectedResult: () =>
      calculateExcelFunction.IF(
        getVal(1, 5) > 0,
        "Ada Diskon",
        "Tidak Ada Diskon"
      ), // F2 > 0
    type: "string",
  },
  {
    level: 2,
    question:
      "Apa **'ID Produk'** untuk 'Laptop XYZ'? (Rumus di sel M2: C2&\"-\"&A2)",
    resultCell: "M2",
    expectedResult: () =>
      calculateExcelFunction.CONCAT(getVal(1, 2), "-", getVal(1, 0)), // C2 & "-" & A2
    type: "string",
  },
  // Level 3: Pencarian & Tanggal (Sekarang jadi Nomor Soal berurutan)
  {
    level: 3,
    question:
      "Apa **'Kode Kategori'** untuk produk 'Laptop XYZ' jika dicari dari Sheet2? (Rumus di sel L2: VLOOKUP(C2,Sheet2!$A$1:$B$4,1,FALSE))",
    resultCell: "L2",
    expectedResult: () =>
      calculateExcelFunction.VLOOKUP(
        getVal(1, 2),
        "Sheet2!$A$1:$B$4",
        1,
        false
      ), // C2
    type: "string",
  },
  {
    level: 3,
    question:
      "Berapa **bulan penjualan** untuk 'Laptop XYZ'? (Rumus di sel N2: MONTH(G2))",
    resultCell: "N2",
    expectedResult: () => calculateExcelFunction.MONTH(getVal(1, 6)), // G2
    type: "number",
  },
  {
    level: 3,
    question:
      "Berapa **'Usia Data (Hari)'** untuk 'Laptop XYZ' hingga hari ini? (Rumus di sel O2: TODAY()-G2)",
    resultCell: "O2",
    expectedResult: () => {
      const startDate = calculateExcelFunction.DATE_PARSE(getVal(1, 6)); // Tanggal penjualan Laptop XYZ (G2)
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
const currentLevelSpan = document.getElementById("current-level"); // ID ini tetap bisa dipakai, kita hanya ganti teksnya.

// --- FUNGSI UTAMA GAME ---

function loadQuestion() {
  if (currentQuestionIndex < questions.length) {
    const q = questions[currentQuestionIndex];
    // Mengubah "Level ${q.level}:" menjadi "Soal ${currentQuestionIndex + 1}:"
    questionTextElement.innerHTML = `**Soal ${currentQuestionIndex + 1}:** ${
      q.question
    }`;
    userResultInput.value = "";
    feedbackElement.textContent = "";
    feedbackElement.className = "feedback";
    nextBtn.style.display = "none";
    checkBtn.style.display = "inline-block";
    // Mengubah "Level: X" menjadi "Nomor Soal: X"
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

  // Normalisasi input dan expectedResult berdasarkan tipe data
  let isCorrect = false;
  if (currentQuestion.type === "number") {
    const parsedUserInput = parseFloat(
      userInputResult.replace(/\./g, "").replace(/,/g, ".")
    );
    if (!isNaN(parsedUserInput)) {
      // Bandingkan angka dengan toleransi kecil untuk floating point
      isCorrect = Math.abs(parsedUserInput - actualExpectedResult) < 0.01;
    }
  } else if (currentQuestion.type === "string") {
    isCorrect =
      userInputResult.toLowerCase() === actualExpectedResult.toLowerCase();
  } else {
    // Fallback untuk tipe lain jika ada
    isCorrect = userInputResult == actualExpectedResult;
  }

  if (isCorrect) {
    feedbackElement.className = "feedback correct";
    feedbackElement.textContent = "Benar! Hasil Anda tepat.";
    nextBtn.style.display = "inline-block";
    checkBtn.style.display = "none";
  } else {
    feedbackElement.className = "feedback incorrect";
    // Tampilkan juga jawaban yang benar untuk pembelajaran
    feedbackElement.innerHTML = `
          Salah. Hasil Anda <code>${userInputResult}</code> tidak sesuai harapan. <br>
          Hasil yang benar adalah <code>${
            typeof actualExpectedResult === "number"
              ? actualExpectedResult.toLocaleString("id-ID")
              : actualExpectedResult
          }</code>. Coba lagi!
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
