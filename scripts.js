let workbook = null;
let worksheet = null;
let employees = [];
let dateColumnIndex = null;
let recognition = null;
let isListening = false;
let currentFilter = "all";
let bulkModeActive = false;
let selectedEmployees = new Set();
let history = [];
let maxHistorySize = 20;
let continuousVoiceMode = false;
let voicePermissionGranted = false;

// متغيرات الموسيقى
let isPlaying = false;
let currentMusicTrack = null;
const musicTracks = {
  1: {
    name: "موجات المحيط الهادئة",
    url: "https://assets.mixkit.co/active_storage/sfx/2338/2338-preview.mp3",
  },
  2: {
    name: "صوت المطر المهدئ",
    url: "https://assets.mixkit.co/active_storage/sfx/2341/2341-preview.mp3",
  },
  3: {
    name: "أصوات الغابة والطيور",
    url: "https://assets.mixkit.co/active_storage/sfx/2340/2340-preview.mp3",
  },
  4: {
    name: "بيانو كلاسيكي هادئ",
    url: "https://assets.mixkit.co/active_storage/music/27/27-preview.mp3",
  },
  5: {
    name: "موسيقى التأمل والاسترخاء",
    url: "https://assets.mixkit.co/active_storage/music/28/28-preview.mp3",
  },
  6: {
    name: "موسيقى الكافيه الهادئة",
    url: "https://assets.mixkit.co/active_storage/music/29/29-preview.mp3",
  },
};

// وظائف مشغل الموسيقى
function toggleMusicPlayer() {
  const player = document.getElementById("musicPlayer");
  if (player.style.display === "none" || player.style.display === "") {
    player.style.display = "block";
  } else {
    player.style.display = "none";
    stopMusic();
  }
}

function changeMusic() {
  const select = document.getElementById("musicSelect");
  const trackId = select.value;

  if (!trackId) {
    stopMusic();
    return;
  }

  const audio = document.getElementById("audioPlayer");
  const track = musicTracks[trackId];

  if (track) {
    audio.src = track.url;
    currentMusicTrack = trackId;
    audio.loop = true;
    updateMusicStatus(`تم اختيار: ${track.name}`);
    playMusic();
  }
}

function playMusic() {
  const audio = document.getElementById("audioPlayer");
  if (audio.src) {
    audio
      .play()
      .then(() => {
        isPlaying = true;
        document.getElementById("playPauseBtn").textContent = "⏸️ إيقاف";
        updateMusicStatus("🎵 موسيقى قيد التشغيل...");
      })
      .catch((err) => {
        updateMusicStatus("❌ لم يتمكن من تشغيل الموسيقى");
      });
  }
}

function stopMusic() {
  const audio = document.getElementById("audioPlayer");
  audio.pause();
  audio.currentTime = 0;
  isPlaying = false;
  document.getElementById("playPauseBtn").textContent = "▶️ تشغيل";
  document.getElementById("musicSelect").value = "";
  updateMusicStatus("اختر موسيقى للبدء");
}

function togglePlayPause() {
  const select = document.getElementById("musicSelect");
  if (!select.value) {
    updateMusicStatus("⚠️ اختر موسيقى أولاً");
    return;
  }

  if (isPlaying) {
    document.getElementById("audioPlayer").pause();
    isPlaying = false;
    document.getElementById("playPauseBtn").textContent = "▶️ تشغيل";
    updateMusicStatus("⏸️ موسيقى معلقة");
  } else {
    playMusic();
  }
}

function setVolume() {
  const slider = document.getElementById("volumeSlider");
  const audio = document.getElementById("audioPlayer");
  const volume = slider.value / 100;
  audio.volume = volume;
  document.getElementById("volumeValue").textContent = slider.value + "%";
}

function updateMusicStatus(message) {
  document.getElementById("musicStatus").textContent = message;
}

function toggleContinuousVoice() {
  continuousVoiceMode = !continuousVoiceMode;
  const btn = document.getElementById("continuousVoiceBtn");

  if (continuousVoiceMode) {
    btn.textContent = "⏸️ إيقاف البحث المستمر";
    btn.style.background = "#28a745";
    document.getElementById("voiceStatus").innerHTML =
      "🎤 البحث المستمر مفعّل - تكلم في أي وقت";
    startContinuousListening();
  } else {
    btn.textContent = "🎤 تشغيل البحث المستمر";
    btn.style.background = "#ff6b6b";
    if (recognition && isListening) {
      recognition.stop();
    }
    document.getElementById("voiceStatus").innerHTML = "";
  }
}

// البحث المستمر
function startContinuousListening() {
  if (!continuousVoiceMode || !recognition) return;

  try {
    recognition.start();
  } catch (e) {
    // إذا كان يعمل بالفعل، انتظر قليلاً
    setTimeout(startContinuousListening, 1000);
  }
}

// تحديث إعداد التعرف على الصوت لدعم الوضع المستمر
function initSpeechRecognition() {
  if ("webkitSpeechRecognition" in window || "SpeechRecognition" in window) {
    const SpeechRecognition =
      window.SpeechRecognition || window.webkitSpeechRecognition;
    recognition = new SpeechRecognition();
    recognition.lang = "ar-SA";
    recognition.continuous = true;
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    recognition.onstart = function () {
      isListening = true;
      voicePermissionGranted = true;
      document.getElementById("voiceBtn").classList.add("listening");
      if (!continuousVoiceMode) {
        document.getElementById("voiceStatus").innerHTML = "🎤 استمع الآن...";
      }
      // إظهار زر البحث المستمر بعد نجاح أول محاولة
      document.getElementById("continuousVoiceBtn").style.display =
        "inline-block";
    };

    recognition.onresult = function (event) {
      const transcript = event.results[0][0].transcript.trim();
      document.getElementById("searchInput").value = transcript;
      displayEmployees(transcript);

      if (continuousVoiceMode) {
        document.getElementById("voiceStatus").innerHTML =
          `✅ "${transcript}" - في انتظار الأمر التالي...`;
      } else {
        document.getElementById("voiceStatus").innerHTML = `✅ "${transcript}"`;
        setTimeout(() => {
          document.getElementById("voiceStatus").innerHTML = "";
        }, 3000);
      }
    };

    recognition.onerror = function (event) {
      isListening = false;
      document.getElementById("voiceBtn").classList.remove("listening");

      let errorMsg = "";
      if (event.error === "no-speech") {
        if (!continuousVoiceMode) {
          errorMsg = "❌ لم أسمع شيئاً";
        }
      } else if (
        event.error === "not-allowed" ||
        event.error === "permission-denied"
      ) {
        errorMsg = "❌ يرجى السماح باستخدام الميكروفون من إعدادات المتصفح";
        document.getElementById("voiceBtn").style.display = "none";
        continuousVoiceMode = false;
      } else if (event.error === "aborted") {
        return;
      } else if (event.error === "network") {
        errorMsg = "❌ مشكلة في الاتصال بالإنترنت";
      }

      if (errorMsg) {
        document.getElementById("voiceStatus").innerHTML = errorMsg;
        if (!continuousVoiceMode) {
          setTimeout(() => {
            document.getElementById("voiceStatus").innerHTML = "";
          }, 3000);
        }
      }
    };

    recognition.onend = function () {
      isListening = false;
      document.getElementById("voiceBtn").classList.remove("listening");

      // في الوضع المستمر، ابدأ من جديد
      if (continuousVoiceMode) {
        setTimeout(startContinuousListening, 300);
      }
    };

    return true;
  }
  return false;
}

// تهيئة البحث الصوتي عند تحميل الصفحة
const speechSupported = initSpeechRecognition();
if (!speechSupported) {
  document.getElementById("voiceBtn").style.display = "none";
  console.log("البحث الصوتي غير مدعوم في هذا المتصفح");
}

function startVoiceSearch() {
  if (!recognition) {
    alert(
      "⚠️ البحث الصوتي غير مدعوم في متصفحك.\nجرب Google Chrome أو Microsoft Edge",
    );
    return;
  }

  if (isListening) {
    // إيقاف الاستماع
    try {
      recognition.stop();
    } catch (e) {
      console.log("تم إيقاف الاستماع");
    }
  } else {
    // بدء الاستماع
    try {
      recognition.start();
    } catch (e) {
      if (e.name === "InvalidStateError") {
        // إذا كان قيد التشغيل، أوقفه ثم ابدأ من جديد
        recognition.stop();
        setTimeout(() => {
          recognition.start();
        }, 100);
      } else {
        console.error("خطأ في بدء البحث الصوتي:", e);
        document.getElementById("voiceStatus").innerHTML =
          "❌ حدث خطأ، حاول مرة أخرى";
      }
    }
  }
}

function markAllAbsent() {
  if (confirm("هل تريد تعيين جميع العمال كغائبين؟")) {
    saveToHistory();
    employees.forEach((emp) => (emp.status = "absent"));
    displayEmployees(document.getElementById("searchInput").value);
    updateStats();
  }
}

function markAllPresent() {
  if (confirm("هل تريد تعيين جميع العمال كحاضرين؟")) {
    saveToHistory();
    employees.forEach((emp) => (emp.status = "present"));
    displayEmployees(document.getElementById("searchInput").value);
    updateStats();
  }
}

function clearAll() {
  if (confirm("هل تريد مسح جميع التسجيلات؟")) {
    saveToHistory();
    employees.forEach((emp) => (emp.status = null));
    displayEmployees(document.getElementById("searchInput").value);
    updateStats();
  }
}

// حفظ الحالة في السجل للتراجع
function saveToHistory() {
  const state = employees.map((emp) => ({ ...emp }));
  history.push(state);
  if (history.length > maxHistorySize) {
    history.shift();
  }
  document.getElementById("undoBtn").style.display =
    history.length > 0 ? "block" : "none";
}

// التراجع عن آخر عملية
function undoLastAction() {
  if (history.length === 0) return;

  const previousState = history.pop();
  employees = previousState.map((emp) => ({ ...emp }));
  displayEmployees(document.getElementById("searchInput").value);
  updateStats();

  if (history.length === 0) {
    document.getElementById("undoBtn").style.display = "none";
  }
}

// فلترة الموظفين
function filterEmployees(filter) {
  currentFilter = filter;

  // تحديث أزرار الفلتر
  document.querySelectorAll(".filter-btn").forEach((btn) => {
    btn.classList.remove("active");
  });
  event.target.classList.add("active");

  displayEmployees(document.getElementById("searchInput").value);
}

// وضع الاختيار المتعدد
function toggleBulkMode() {
  bulkModeActive = !bulkModeActive;
  selectedEmployees.clear();

  if (bulkModeActive) {
    event.target.textContent = "✓ إلغاء الاختيار المتعدد";
    event.target.style.background = "#dc3545";
    alert(
      "الآن يمكنك اختيار عدة عمال بالنقر عليهم، ثم استخدم الأزرار لتغيير حالتهم جميعاً",
    );
  } else {
    event.target.textContent = "📦 وضع الاختيار المتعدد";
    event.target.style.background = "#17a2b8";
  }

  displayEmployees(document.getElementById("searchInput").value);
}

// تصدير إلى CSV
function exportToCSV() {
  const selectedDate = document.getElementById("attendanceDate").value;
  let csv = "الاسم,الحالة,التاريخ\n";

  employees.forEach((emp) => {
    const status =
      emp.status === "present"
        ? "حاضر"
        : emp.status === "absent"
          ? "غائب"
          : "غير محدد";
    csv += `"${emp.name}","${status}","${selectedDate}"\n`;
  });

  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `حضور_${selectedDate}.csv`;
  link.click();
}

// طباعة التقرير
function printReport() {
  window.print();
}

// اختصارات لوحة المفاتيح
document.addEventListener("keydown", function (e) {
  // Ctrl + S للحفظ
  if (e.ctrlKey && e.key === "s") {
    e.preventDefault();
    saveAttendance();
  }
  // Ctrl + F للبحث
  else if (e.ctrlKey && e.key === "f") {
    e.preventDefault();
    document.getElementById("searchInput").focus();
  }
  // Ctrl + Z للتراجع
  else if (e.ctrlKey && e.key === "z") {
    e.preventDefault();
    undoLastAction();
  }
  // Ctrl + P للطباعة
  else if (e.ctrlKey && e.key === "p") {
    e.preventDefault();
    printReport();
  }
  // Ctrl + A للجميع حاضر
  else if (e.ctrlKey && e.key === "a") {
    e.preventDefault();
    markAllPresent();
  }
  // Ctrl + D للجميع غائب
  else if (e.ctrlKey && e.key === "d") {
    e.preventDefault();
    markAllAbsent();
  }
  // ? لإظهار/إخفاء الاختصارات
  else if (e.key === "?") {
    const panel = document.getElementById("shortcutsPanel");
    panel.style.display = panel.style.display === "none" ? "block" : "none";
  }
});

// تعيين التاريخ الحالي
document.getElementById("attendanceDate").valueAsDate = new Date();

// تحميل ملف الإكسل
document.getElementById("excelFile").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array" });
      worksheet = workbook.Sheets[workbook.SheetNames[0]];

      loadEmployees();
      document.getElementById("fileInfo").innerHTML =
        '<div class="message success">✅ تم تحميل الملف بنجاح!</div>';
    } catch (error) {
      document.getElementById("fileInfo").innerHTML =
        '<div class="message warning">❌ خطأ في قراءة الملف</div>';
    }
  };
  reader.readAsArrayBuffer(file);
});

function loadEmployees() {
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  employees = [];

  // البحث عن عمود التاريخ المحدد
  const selectedDate = document.getElementById("attendanceDate").value;
  const headers = jsonData[0];

  dateColumnIndex = null;
  if (selectedDate) {
    const searchDate = new Date(selectedDate);
    for (let i = 1; i < headers.length; i++) {
      const headerDate = XLSX.SSF.parse_date_code(headers[i]);
      if (headerDate) {
        const cellDate = new Date(headerDate.y, headerDate.m - 1, headerDate.d);
        if (cellDate.toDateString() === searchDate.toDateString()) {
          dateColumnIndex = i;
          break;
        }
      }
    }
  }

  // تحميل الأسماء
  for (let i = 1; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (row[0]) {
      const name = String(row[0]).trim();
      let status = null;

      if (dateColumnIndex !== null && row[dateColumnIndex] !== undefined) {
        status =
          row[dateColumnIndex] === 1
            ? "present"
            : row[dateColumnIndex] === 0
              ? "absent"
              : null;
      }

      // إذا كان الخيار مفعل وليس هناك قيمة محفوظة، اجعله غائب
      if (
        document.getElementById("startWithAbsent").checked &&
        status === null
      ) {
        status = "absent";
      }

      employees.push({
        name: name,
        rowIndex: i,
        status: status,
      });
    }
  }

  displayEmployees();
  updateStats();
  document.getElementById("quickActionsSection").style.display = "block";
}

function displayEmployees(filter = "") {
  const grid = document.getElementById("employeeGrid");
  grid.innerHTML = "";

  let filteredEmployees = filter
    ? employees.filter((emp) => emp.name.includes(filter))
    : employees;

  // تطبيق فلتر الحالة
  if (currentFilter === "present") {
    filteredEmployees = filteredEmployees.filter(
      (emp) => emp.status === "present",
    );
  } else if (currentFilter === "absent") {
    filteredEmployees = filteredEmployees.filter(
      (emp) => emp.status === "absent",
    );
  } else if (currentFilter === "unmarked") {
    filteredEmployees = filteredEmployees.filter((emp) => emp.status === null);
  }

  filteredEmployees.forEach((emp, index) => {
    const card = document.createElement("div");
    card.className = "employee-card";
    if (emp.status === "present") card.classList.add("present");
    if (emp.status === "absent") card.classList.add("absent");
    if (bulkModeActive) card.classList.add("bulk-mode");
    if (selectedEmployees.has(emp.name))
      card.classList.add("selected-for-bulk");

    const encodedName = emp.name.replace(/'/g, "&#39;").replace(/"/g, "&quot;");
    card.innerHTML = `
                    <div class="employee-name">${encodedName}</div>
                    <div class="employee-status">
                        <button class="status-btn present-btn" data-name="${encodedName}" onclick="markPresent(this.dataset.name); event.stopPropagation();">
                            ✓ حاضر
                        </button>
                        <button class="status-btn absent-btn" data-name="${encodedName}" onclick="markAbsent(this.dataset.name); event.stopPropagation();">
                            ✗ غائب
                        </button>
                    </div>
                `;

    // نقرة على البطاقة نفسها
    card.onclick = function () {
      if (bulkModeActive) {
        // في وضع الاختيار المتعدد
        if (selectedEmployees.has(emp.name)) {
          selectedEmployees.delete(emp.name);
        } else {
          selectedEmployees.add(emp.name);
        }
        displayEmployees(document.getElementById("searchInput").value);
      } else {
        // التبديل العادي
        if (emp.status === "absent" || emp.status === null) {
          markPresent(emp.name);
        } else {
          markAbsent(emp.name);
        }
      }
    };

    grid.appendChild(card);
  });

  // عرض الأزرار للاختيار المتعدد
  if (bulkModeActive && selectedEmployees.size > 0) {
    const bulkActions = document.createElement("div");
    bulkActions.className = "action-buttons";
    bulkActions.innerHTML = `
                    <button class="action-btn" style="background: #28a745;" onclick="markSelectedPresent()">
                        ✓ تعيين المحددين كحاضر (${selectedEmployees.size})
                    </button>
                    <button class="action-btn" style="background: #dc3545;" onclick="markSelectedAbsent()">
                        ✗ تعيين المحددين كغائب (${selectedEmployees.size})
                    </button>
                `;
    grid.appendChild(bulkActions);
  }

  document.getElementById("statsSection").style.display = "block";
  document.getElementById("searchSection").style.display = "block";
  document.getElementById("employeeSection").style.display = "block";
}

// تعيين المحددين كحاضر
function markSelectedPresent() {
  saveToHistory();
  selectedEmployees.forEach((name) => {
    const emp = employees.find((e) => e.name === name);
    if (emp) emp.status = "present";
  });
  selectedEmployees.clear();
  displayEmployees(document.getElementById("searchInput").value);
  updateStats();
}

// تعيين المحددين كغائب
function markSelectedAbsent() {
  saveToHistory();
  selectedEmployees.forEach((name) => {
    const emp = employees.find((e) => e.name === name);
    if (emp) emp.status = "absent";
  });
  selectedEmployees.clear();
  displayEmployees(document.getElementById("searchInput").value);
  updateStats();
}

function markPresent(name) {
  saveToHistory();
  const emp = employees.find((e) => e.name === name);
  if (emp) {
    emp.status = "present";
    displayEmployees(document.getElementById("searchInput").value);
    updateStats();
  }
}

function markAbsent(name) {
  saveToHistory();
  const emp = employees.find((e) => e.name === name);
  if (emp) {
    emp.status = "absent";
    displayEmployees(document.getElementById("searchInput").value);
    updateStats();
  }
}

function updateStats() {
  const total = employees.length;
  const present = employees.filter((e) => e.status === "present").length;
  const absent = employees.filter((e) => e.status === "absent").length;
  const unmarked = total - present - absent;

  document.getElementById("totalCount").textContent = total;
  document.getElementById("presentCount").textContent = present;
  document.getElementById("absentCount").textContent = absent;
  document.getElementById("unmarkedCount").textContent = unmarked;
}

function saveAttendance() {
  if (!workbook || dateColumnIndex === null) {
    alert("⚠️ الرجاء اختيار تاريخ صحيح");
    return;
  }

  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  // تحديث البيانات
  employees.forEach((emp) => {
    if (emp.status) {
      const value = emp.status === "present" ? 1 : 0;
      jsonData[emp.rowIndex][dateColumnIndex] = value;
    }
  });

  // إنشاء ملف جديد
  const newWorksheet = XLSX.utils.aoa_to_sheet(jsonData);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "ورقة1");

  // تحميل الملف
  const selectedDate = document.getElementById("attendanceDate").value;
  const fileName = `حضور_وغياب_${selectedDate}.xlsx`;
  XLSX.writeFile(newWorkbook, fileName);

  alert("✅ تم حفظ الملف بنجاح!");
}

function resetAll() {
  if (confirm("هل أنت متأكد من إعادة تعيين جميع البيانات؟")) {
    employees.forEach((emp) => (emp.status = null));
    displayEmployees();
    updateStats();
  }
}

// البحث
// البحث
document.getElementById("searchInput").addEventListener("input", function (e) {
  displayEmployees(e.target.value);
});

// تغيير التاريخ
document
  .getElementById("attendanceDate")
  .addEventListener("change", function () {
    if (workbook) {
      loadEmployees();
    }
  });

// تهيئة مشغل الموسيقى
document.addEventListener("DOMContentLoaded", function () {
  const player = document.getElementById("musicPlayer");
  if (player) {
    player.style.display = "none";
  }
  setVolume();
});
