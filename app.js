/**
 * app.js - Client-side Logic for Annual Health Check Registration System
 */

// Global Configuration
const CONFIG = {
  // Paste your deployed Google Sheets Web App URL here
  apiUrl: "https://script.google.com/macros/s/AKfycbxl0kN1jKq-u5Wf--ftanFX0hSbrp7tPgdM9c149c0kUs3bP3ggdKrnOTOq0tgf1TaaDA/exec",
  currentMode: "api", // Defaults to "api" to connect directly to the user's live sheet
};

// Global State
const STATE = {
  activeEmployee: null,
  activeRegistration: null,
  configDates: [],
  configTimeSlots: [],
  registrationCounts: {},
  selectedTimeSlot: "", // Holds the currently clicked slot text
  catalogCategory: "all",
  catalogSearchQuery: "",
  isAdminAuthenticated: false,
  adminSubTab: "slots",
  adminDashboardData: null
};

// --- Special Self-Pay Test Catalog Data ---
const SPECIAL_TESTS = [
  { id: 1, name: "น้ำตาลในเลือด (FBS)", purpose: "ดูค่าน้ำตาลในเลือดประเมินดูภาวะเบาหวานเบื้องต้น", price: 30, fasting: true, category: "blood" },
  { id: 2, name: "ตรวจหาระดับไขมันในเลือด (Cholesterol, Triglyceride, HDL, LDL)", purpose: "ดูระดับไขมันแต่ละชนิดในเลือด", price: 160, fasting: true, category: "blood", notes: "ฟรีสำหรับพนักงานที่อายุ 35 ปีขึ้นไป" },
  { id: 3, name: "ตรวจสมรรถภาพการทำงานของตับ (SGOT, SGPT, ALK)", purpose: "ดูความผิดปกติของตับ ที่อาจทำให้เป็นโรคตับอักเสบ โรคมะเร็งตับ ผู้ที่มีประวัติดื่มสุราเป็นประจำควรตรวจเป็นประจำทุกปี", price: 150, fasting: false, category: "blood" },
  { id: 4, name: "ตรวจการทำงานของไต (BUN, Creatinine)", purpose: "ตรวจสอบการทำงานของไต ว่าอยู่ปกติหรือไม่", price: 60, fasting: false, category: "blood" },
  { id: 5, name: "ตรวจภาวะไทรอยด์ (TFT = FT3, FT4, TSH)", purpose: "ไทรอยด์ฮอร์โมน ช่วยควบคุมการเผาผลาญของร่างกาย ช่วยบ่งบอกว่าต่อมไทรอยด์ทำงานผิดปกติหรือไม่", price: 400, fasting: false, category: "blood" },
  { id: 6, name: "ตรวจหาภาวะโรคเก๊าท์ (Uric Acid)", purpose: "ดูปริมาณของกรดยูริกในเลือดเพื่อวินิจฉัยโรคเก๊าท์", price: 50, fasting: false, category: "blood" },
  { id: 7, name: "ตรวจหาน้ำตาลสะสม (คัดกรองโรคเบาหวาน) (HbA1C)", purpose: "วัดระดับน้ำตาลสะสมในเลือดตลอดระยะเวลา 4 เดือนที่ผ่านมา", price: 300, fasting: false, category: "blood" },
  { id: 8, name: "ตรวจเกลือแร่ในเลือด (Electrolyte)", purpose: "ตรวจปริมาณ โซเดียม โปแตสเซียม ที่เหมาะสมในร่างกาย ซึ่งมีผลต่อการทำงานของกล้ามเนื้อหัวใจและกล้ามเนื้อส่วนอื่น", price: 300, fasting: false, category: "blood" },
  { id: 9, name: "ตรวจหาระดับแคลเซียมในเลือด (Total Calcium)", purpose: "ดูปริมาณแคลเซียมในเลือด", price: 50, fasting: false, category: "blood" },
  { id: 10, name: "ตรวจหาเชื้อไวรัสตับอักเสบชนิดบี (HBs Ag)", purpose: "ตรวจว่ามีเชื้อตับอักเสบ B อยู่ในร่างกายหรือไม่", price: 100, fasting: false, category: "blood" },
  { id: 11, name: "ตรวจหาภูมิคุ้มกันเชื้อไวรัสตับอักเสบบี (Anti-HBs)", purpose: "ตรวจเพื่อดูว่าขณะนี้มีภูมิต้านทานเชื้อไวรัสตับอักเสบชนิด B หรือไม่", price: 150, fasting: false, category: "blood" },
  { id: 12, name: "ตรวจหาไวรัสตับอักเสบซี (Anti-HCV)", purpose: "คือการตรวจหาเชื้อตับอักเสบ C อยู่ในร่างกายหรือไม่", price: 300, fasting: false, category: "blood" },
  { id: 13, name: "ตรวจหากรุ๊ปเลือด (Blood Group)", purpose: "ตรวจระบุหมู่โลหิต (Blood Group)", price: 50, fasting: false, category: "blood" },
  { id: 14, name: "ตรวจคลื่นไฟฟ้าหัวใจ (EKG)", purpose: "คือการตรวจความสมบูรณ์ของการทำงานไฟฟ้าหัวใจ", price: 200, fasting: false, category: "blood", notes: "LPN1: ตรวจที่ห้องพยาบาล Plant 3 (16, 30 ต.ค. 68) | LPN2: ห้องพยาบาล (17, 31 ต.ค. 68)" },
  { id: 15, name: "ตรวจคัดกรองมะเร็งต่อมลูกหมาก (PSA)", purpose: "การตรวจสารบ่งชี้มะเร็งต่อมลูกหมากในผู้ชาย", price: 300, fasting: false, category: "cancer", gender: "M" },
  { id: 16, name: "ตรวจคัดกรองมะเร็งตับ (AFP)", purpose: "การตรวจสารบ่งชี้มะเร็งเพื่อช่วยวินิจฉัยมะเร็งตับ", price: 300, fasting: false, category: "cancer" },
  { id: 17, name: "ตรวจคัดกรองมะเร็งทางเดินอาหาร/มะเร็งลำไส้ (CEA)", purpose: "การตรวจสารบ่งชี้มะเร็งเพื่อช่วยวินิจฉัยมะเร็งทางเดินอาหาร", price: 300, fasting: false, category: "cancer" },
  { id: 18, name: "ตรวจคัดกรองมะเร็งรังไข่ (CA125)", purpose: "การตรวจสารบ่งชี้มะเร็งรังไข่ ส่วนใหญ่ในผู้หญิง", price: 500, fasting: false, category: "cancer", gender: "F" },
  { id: 19, name: "ตรวจคัดกรองมะเร็งตับอ่อน (CA 19-9)", purpose: "การตรวจสารบ่งชี้มะเร็งตับอ่อนและมะเร็งของท่อน้ำดี", price: 500, fasting: false, category: "cancer" },
  { id: 20, name: "ตรวจคัดกรองมะเร็งเต้านม (CA153)", purpose: "การตรวจสารบ่งชี้มะเร็งเต้านมส่วนใหญ่ในผู้หญิง", price: 500, fasting: false, category: "cancer", gender: "F" },
  { id: 21, name: "ตรวจหาร่องรอยเชื้อไวรัสเอดส์ (Anti-HIV)", purpose: "ตรวจหาร่องรอยเชื้อไวรัสเอดส์ (Anti HIV)", price: 180, fasting: false, category: "cancer" },
  { id: 22, name: "ตรวจคัดกรองโรคธาลัสซีเมีย (HB Typing)", purpose: "การตรวจคัดกรองโรคเลือดจางธาลัสซีเมีย", price: 600, fasting: false, category: "cancer" },
  { id: 23, name: "ตรวจสมรรถภาพการมองเห็น", purpose: "เป็นการตรวจที่ช่วยประเมินการมองเห็นชัดเจนเพียงใด", price: 30, fasting: false, category: "blood" },
  { id: 24, name: "ตรวจสมรรถภาพปอด", purpose: "เป็นการตรวจการทำงานของปอด", price: 30, fasting: false, category: "blood" },
  { id: 25, name: "ตรวจสมรรถภาพการได้ยิน", purpose: "เป็นการตรวจการทำงานของหู", price: 30, fasting: false, category: "blood", location: "ตรวจที่รถ X-ray" },
  { id: 26, name: "ตรวจหาปริมาณภูมิหลังฉีดวัคซีน Covid-19", purpose: "เป็นการตรวจหาปริมาณภูมิคุ้มกันโรคโควิด-19 ในร่างกาย", price: 800, fasting: false, category: "blood" },
  
  // Welfare program items (items 27-33)
  { id: 27, name: "Mammogram", purpose: "การตรวจหาความผิดปกติของเต้านม (ผู้หญิง)", price: 1800, fasting: false, category: "cancer", gender: "F", welfare: { limit: 1500, minAge: 50 } },
  { id: 28, name: "Thin Prep", purpose: "การตรวจคัดกรองมะเร็งปากมดลูก (ผู้หญิง)", price: 1200, fasting: false, category: "cancer", gender: "F", welfare: { limit: 1500, minAge: 50 } },
  { id: 29, name: "Prostate Screening", purpose: "การตรวจมะเร็งต่อมลูกหมากโดยแพทย์ (ผู้ชาย)", price: 300, fasting: false, category: "cancer", gender: "M", welfare: { limit: 1000, minAge: 50 } },
  { id: 30, name: "CT Calcium Score", purpose: "ตรวจแคลเซียมที่ผนังหลอดเลือดหัวใจ", price: 5500, fasting: false, category: "scan", welfare: { limit: 2000, minAge: 50, minLevel: "M4/T5" } },
  { id: 31, name: "Colonoscopy / Colonoscopy + Biopsy", purpose: "การส่องกล้องตรวจลำไส้ใหญ่ (14,000 ฿) หรือส่องกล้องร่วมกับตัดชิ้นเนื้อ (15,000 ฿)", price: 14000, fasting: false, category: "scan", welfare: { limit: 15000, minAge: 50, minLevel: "M5/T6", isColon: true } },
  { id: 33, name: "MRI Brain (MRI & MRA)", purpose: "การตรวจสมองด้วยคลื่นแม่เหล็กไฟฟ้า (MRI Brain 24,000 ฿ / MRI+MRA 29,000 ฿)", price: 24000, fasting: false, category: "scan", welfare: { limit: 10000, minAge: 50, minLevel: "M5/T6", isMri: true } }
];

// --- Test Programs Data ---
const PROGRAM_TESTS = {
  "โปรแกรม MGR": [
    { name: "พบแพทย์ (PE)", npo: false },
    { name: "สายตา (EYE)", npo: false },
    { name: "ความสมบูรณ์ของเม็ดเลือด (CBC)", npo: false },
    { name: "ปัสสาวะ (UA)", npo: false },
    { name: "เอกซเรย์ (X-RAY)", npo: false },
    { name: "ไขมัน (Cholesterol, TG, LDL, HDL)", npo: true },
    { name: "น้ำตาล (FBS)", npo: true },
    { name: "การทำงานของตับ (SGOT, SGPT)", npo: false },
    { name: "การทำงานของไต (BUN, Cr, eGFR)", npo: false },
    { name: "กรดยูริค (เก๊าท์)", npo: false },
    { name: "ตรวจหาน้ำตาลสะสม (คัดกรองโรคเบาหวาน) (HbA1C)", npo: false }
  ],
  "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป": [
    { name: "พบแพทย์ (PE)", npo: false },
    { name: "สายตา (EYE)", npo: false },
    { name: "เจาะเลือด CBC", npo: false },
    { name: "ปัสสาวะ (UA)", npo: false },
    { name: "เอกซเรย์ (X-RAY)", npo: false },
    { name: "ไขมัน (Cholesterol, TG, LDL, HDL)", npo: true }
  ],
  "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี": [
    { name: "พบแพทย์ (PE)", npo: false },
    { name: "สายตา (EYE)", npo: false },
    { name: "เจาะเลือด CBC", npo: false },
    { name: "ปัสสาวะ (UA)", npo: false },
    { name: "เอกซเรย์ (X-RAY)", npo: false }
  ]
};

// --- Mock Data for Offline Mode ---
const MOCK_EMPLOYEES = [
  { employeeId: "003049", firstName: "วิชัย", lastName: "สุขประเสริฐกุล", department: "OPT", defaultLocation: "LPN1", programName: "โปรแกรม MGR", age: 59, gender: "M", programGroup: "โปรแกรม MGR", riskProgram: "" },
  { employeeId: "004148", firstName: "ประภาพร", lastName: "ศรีประดู่", department: "HRDS", defaultLocation: "LPN2", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 57, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "ตรวจการได้ยิน (Audiogram), ตรวจปัสสาวะหาสารเคมี", isPregnant: true },
  { employeeId: "004379", firstName: "ประคอง", lastName: "อ้อยงาม", department: "QM", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 54, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "ตรวจหาสารตะกั่ว" },
  { employeeId: "004766", firstName: "พวงเพชร", lastName: "มณีฉาย", department: "CADT", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 55, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "" },
  { employeeId: "005933", firstName: "ระเบียบ", lastName: "ปาละรัตน์", department: "QM", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 60, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "" },
  { employeeId: "006078", firstName: "อดิเรก", lastName: "อ่อนพรม", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 57, gender: "M", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "" },
  { employeeId: "006125", firstName: "ยุทธนา", lastName: "สุยะนันทน์", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรม MGR", age: 56, gender: "M", programGroup: "โปรแกรม MGR", riskProgram: "ตรวจคลื่นไฟฟ้าหัวใจ (EKG), ตรวจสมรรถภาพปอด" },
  { employeeId: "007860", firstName: "วิไล", lastName: "แสวงศรี", department: "OP4", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 49, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "" },
  { employeeId: "009892", firstName: "สมชาย", lastName: "ม่วงไหม", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรม MGR", age: 56, gender: "M", programGroup: "โปรแกรม MGR", riskProgram: "" },
  { employeeId: "010268", firstName: "ยอดธง", lastName: "กรวิรัตน์", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรม MGR", age: 55, gender: "M", programGroup: "MGR", riskProgram: "" },
  { employeeId: "011382", firstName: "เพียงอัมพร", lastName: "องค์วิศิษฐ์", department: "TRF", defaultLocation: "LPN1", programName: "โปรแกรม MGR", age: 57, gender: "F", programGroup: "โปรแกรม MGR", riskProgram: "" },
  { employeeId: "011999", firstName: "ณัฐพงษ์", lastName: "รักเรียน", department: "IT", defaultLocation: "LPN1", programName: "โปรแกรมอายุไม่ถึง 35 ปี", age: 28, gender: "M", programGroup: "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี", riskProgram: "", checkupRight: "ไม่มีสิทธิ์ (อายุงานไม่ถึง 6 เดือน)" }
];

const MOCK_CONFIG_DATES = [
  { location: "LPN1", dateString: "30 กันยายน 2569", team: "ทีม A" },
  { location: "LPN1", dateString: "1 ตุลาคม 2569", team: "ทีม A" },
  { location: "LPN1", dateString: "6 ตุลาคม 2569", team: "ทีม B" },
  { location: "LPN1", dateString: "7 ตุลาคม 2569", team: "ทีม B" },
  { location: "LPN2", dateString: "2 ตุลาคม 2569", team: "ทีม A" },
  { location: "LPN2", dateString: "5 ตุลาคม 2569", team: "ทีม B" }
];

const MOCK_CONFIG_TIMESLOTS = [
  { slotTime: "08:00 - 08:30", limit: 50 },
  { slotTime: "08:30 - 09:00", limit: 50 },
  { slotTime: "09:00 - 09:30", limit: 50 },
  { slotTime: "09:30 - 10:00", limit: 50 },
  { slotTime: "10:00 - 10:30", limit: 50 },
  { slotTime: "10:30 - 11:00", limit: 50 },
  { slotTime: "11:00 - 11:30", limit: 50 },
  { slotTime: "11:30 - 12:00", limit: 50 },
  { slotTime: "12:00 - 12:30", limit: 50 },
  { slotTime: "12:30 - 13:00", limit: 50 },
  { slotTime: "13:00 - 13:30", limit: 50 },
  { slotTime: "13:30 - 14:00", limit: 50 }
];

// Initialize Mock Registrations locally in LocalStorage
if (!localStorage.getItem("MOCK_REGISTRATIONS")) {
  const initialRegs = [
    { employeeId: "003049", firstName: "วิชัย", lastName: "สุขประเสริฐกุล", department: "OPT", phone: "2201", shift: "ทีม A", location: "LPN1", dateString: "30 กันยายน 2569", timeString: "07:00 - 07:30", cancerTest: "ตรวจคัดกรองมะเร็งต่อมลูกหมาก (PSA)", timestamp: "2026-07-30 08:30:12" },
    { employeeId: "004148", firstName: "ประภาพร", lastName: "ศรีประดู่", department: "HRDS", phone: "1041", shift: "ทีม B", location: "LPN2", dateString: "5 ตุลาคม 2569", timeString: "09:00 - 09:30", cancerTest: "", timestamp: "2026-07-30 09:12:44" }
  ];
  localStorage.setItem("MOCK_REGISTRATIONS", JSON.stringify(initialRegs));
}

// --- Initialize App ---
document.addEventListener("DOMContentLoaded", () => {
  setupModeSelector();
  loadConfigAndCounts();
  initRealTimeSync();
  validateFormCompletion();
});

function setupModeSelector() {
  const selector = document.getElementById("app-mode-selector");
  selector.value = CONFIG.currentMode;
  updateStatusDot(CONFIG.currentMode);
  
  selector.addEventListener("change", (e) => {
    CONFIG.currentMode = e.target.value;
    updateStatusDot(CONFIG.currentMode);
    showToast(`สลับโหมดเป็น: ${CONFIG.currentMode === "api" ? "เชื่อมต่อ Google Sheet" : "จำลองระบบออฟไลน์ (Mock)"}`, "info");
    
    resetForm();
    document.getElementById("result-card-container").classList.remove("visible");
    loadConfigAndCounts();
  });
}

function updateStatusDot(mode) {
  const dot = document.getElementById("status-dot");
  if (mode === "api") {
    dot.className = "status-dot";
  } else {
    dot.className = "status-dot offline";
  }
}

// --- Load config settings & counts dynamically ---
async function loadConfigAndCounts() {
  if (CONFIG.currentMode === "mock") {
    STATE.configDates = MOCK_CONFIG_DATES;
    STATE.configTimeSlots = MOCK_CONFIG_TIMESLOTS;
    
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    STATE.registrationCounts = {};
    regs.forEach(r => {
      const key = `${r.location}|${r.dateString}|${r.timeString}`;
      STATE.registrationCounts[key] = (STATE.registrationCounts[key] || 0) + 1;
    });
  } else {
    showLoader("กำลังดึงข้อมูลกำหนดการและสิทธิ์การจองล่าสุด...");
    try {
      const response = await callApi("getConfigAndSlots", []);
      if (response && response.success) {
        STATE.configDates = response.data.dates;
        STATE.configTimeSlots = response.data.timeSlots;
        STATE.registrationCounts = response.data.registrationCounts;
      } else {
        throw new Error(response.error || "ดึงข้อมูลล้มเหลว");
      }
    } catch (err) {
      console.error(err);
      showToast("ไม่สามารถเชื่อมต่อ Google Sheets API ได้ จะใช้ข้อมูลจำลองแทนชั่วคราว", "warning");
      STATE.configDates = MOCK_CONFIG_DATES;
      STATE.configTimeSlots = MOCK_CONFIG_TIMESLOTS;
    } finally {
      hideLoader();
    }
  }
}

// --- Navigation Tabs Control ---
function switchTab(tabName) {
  document.querySelectorAll(".tab-btn").forEach(btn => btn.classList.remove("active"));
  document.getElementById(`tab-btn-${tabName}`).classList.add("active");
  
  document.querySelectorAll(".tab-content").forEach(panel => panel.classList.remove("active"));
  document.getElementById(`tab-${tabName}`).classList.add("active");
  
  if (tabName === "special-catalog") {
    renderSpecialCatalog();
  } else if (tabName === "admin") {
    checkAdminState();
  }
}

// --- Lookup Employee Details ---
async function lookupEmployee() {
  const empInput = document.getElementById("reg-emp-id");
  let empId = empInput.value.trim();
  
  if (!empId) {
    showToast("กรุณากรอกรหัสพนักงาน", "warning");
    return;
  }
  
  if (/^\d+$/.test(empId)) {
    empId = empId.padStart(6, '0');
    empInput.value = empId;
  }
  
  showLoader("กำลังค้นหาข้อมูลพนักงาน...");
  
  try {
    if (CONFIG.currentMode === "mock") {
      await new Promise(resolve => setTimeout(resolve, 500));
      const emp = MOCK_EMPLOYEES.find(e => e.employeeId === empId);
      if (emp) {
        const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
        const reg = regs.find(r => r.employeeId === empId);
        handleEmployeeLookupResult(emp, reg);
      } else {
        handleEmployeeLookupResult(null, null);
      }
    } else {
      const res = await callApi("getEmployeeData", [empId]);
      if (res && res.success && res.data) {
        let reg = null;
        try {
          const regRes = await callApi("getRegistrationByEmpId", [empId]);
          if (regRes && regRes.success) {
            reg = regRes.data;
          }
        } catch (e) {
          console.warn("Could not fetch registration:", e);
        }
        handleEmployeeLookupResult(res.data, reg);
      } else {
        throw new Error((res && res.error) || "เกิดข้อผิดพลาดในการดึงข้อมูล");
      }
    }
  } catch (err) {
    console.error(err);
    showToast(`ไม่พบรหัสพนักงาน: ${err.message}`, "error");
    handleEmployeeLookupResult(null, null);
  } finally {
    hideLoader();
  }
}

function handleEmployeeLookupResult(employee, registration = null) {
  const profileBox = document.getElementById("employee-profile-box");
  
  if (!employee) {
    showToast("ไม่พบข้อมูลพนักงานท่านนี้ กรุณาตรวจสอบรหัสพนักงานใหม่อีกครั้ง", "error");
    profileBox.style.display = "none";
    STATE.activeEmployee = null;
    return;
  }
  
  if (employee.checkupRight && employee.checkupRight.indexOf("ยังไม่มีสิทธิ์ตรวจ") !== -1) {
    showToast(`ขออภัย รหัสพนักงานนี้ยังไม่มีสิทธิ์ลงทะเบียนตรวจสุขภาพประจำปี เนื่องจากจากเข้างานยังไม่ครบ 6 เดือน (${employee.checkupRight})`, "error");
    profileBox.style.display = "none";
    STATE.activeEmployee = null;
    return;
  }
  
  STATE.activeEmployee = employee;
  
  // Populate Profile UI (Age completely omitted for PDPA compliance)
  document.getElementById("profile-name").textContent = `${employee.firstName} ${employee.lastName}`;
  document.getElementById("profile-dept").textContent = employee.department;
  
  const programGroup = employee.programGroup;
  document.getElementById("profile-program-group").textContent = programGroup;
  
  // Populate checkup items grid
  renderCheckupList(employee, !!employee.isPregnant);
  
  // Handle Cancer selection for MGR
  const cancerSelectionWrapper = document.getElementById("cancer-selection-wrapper");
  const radioButtons = document.querySelectorAll('input[name="cancerTest"]');
  
  if (programGroup === "โปรแกรม MGR") {
    cancerSelectionWrapper.style.display = "block";
    radioButtons.forEach(btn => btn.required = true);
  } else {
    cancerSelectionWrapper.style.display = "none";
    radioButtons.forEach(btn => {
      btn.required = false;
      btn.checked = false;
    });
    document.querySelectorAll(".cancer-option").forEach(opt => opt.classList.remove("selected"));
  }
  
  const prevRegCard = document.getElementById("prev-reg-card");
  if (registration) {
    prevRegCard.style.display = "flex";
    document.getElementById("prev-reg-loc-time").textContent = 
      `สถานที่: ${registration.location} | วันที่: ${registration.dateString} | เวลา: ${registration.timeString}`;
    
    // Preset form values
    document.getElementById("reg-phone").value = registration.phone || "";
    document.getElementById("reg-shift").value = registration.shift || "";
    document.getElementById("reg-location").value = registration.location || "";
    
    // Trigger populating the date list
    onShiftOrLocationChange();
    
    document.getElementById("reg-date").value = registration.dateString || "";
    
    // Set selected slot state
    STATE.selectedTimeSlot = registration.timeString || "";
    document.getElementById("selected-time-slot").value = registration.timeString || "";
    
    // Render time slots and select the previous one
    renderTimeSlots(true);
    
    // Handle Cancer selection for MGR
    if (programGroup === "โปรแกรม MGR" && registration.cancerTest) {
      const radio = document.querySelector(`input[name="cancerTest"][value="${registration.cancerTest}"]`);
      if (radio) {
        radio.checked = true;
        const label = radio.closest(".cancer-option");
        if (label) selectCancerOption(label);
      }
    }
    
    // SSO consent selection removed
  } else {
    prevRegCard.style.display = "none";
    
    // Try to set default location based on sheet data
    const locSelect = document.getElementById("reg-location");
    if (employee.defaultLocation && (employee.defaultLocation === "LPN1" || employee.defaultLocation === "LPN2")) {
      locSelect.value = employee.defaultLocation;
    } else {
      locSelect.selectedIndex = 0;
    }
    
    // Reset input fields
    document.getElementById("reg-shift").selectedIndex = 0;
    document.getElementById("reg-phone").value = "";
    
    const dateSelect = document.getElementById("reg-date");
    dateSelect.innerHTML = `<option value="" disabled selected>-- กรุณาเลือกสถานที่และกะก่อน --</option>`;
    dateSelect.disabled = true;
    
    const gridContainer = document.getElementById("time-slots-grid-container");
    gridContainer.innerHTML = `<div class="empty-slots-msg">กรุณาเลือกสถานที่ กะทำงาน และวันที่ตรวจสุขภาพด้านบนก่อน</div>`;
    STATE.selectedTimeSlot = "";
    document.getElementById("selected-time-slot").value = "";
    
  }
  
  profileBox.style.display = "block";
  profileBox.scrollIntoView({ behavior: "smooth", block: "nearest" });
  validateFormCompletion();
}

// --- Render checkup list dynamically ---
function renderCheckupList(employee, isPregnant = false) {
  const itemsList = document.getElementById("checkup-items-list");
  if (!itemsList) return;
  itemsList.innerHTML = "";
  
  // Determine if we should show SSO items (always true now that opt-out is removed)
  const useSso = true;
  
  const programGroup = employee.programGroup;
  const baseTests = PROGRAM_TESTS[programGroup] || [];
  
  // Clone base tests
  let tests = baseTests.map(t => ({ ...t, isSso: false }));
  
  const age = parseInt(employee.age || 0, 10);
  const isFemale = employee.gender === "F" || employee.gender === "Female" || !!employee.isPregnant;
  
  if (useSso) {
    // 3. FBS (age >= 35)
    if (age >= 35) {
      let hasFbs = false;
      tests.forEach(t => {
        if (programGroup === "โปรแกรม MGR") return;
        if (t.name.includes("FBS") || t.name.includes("น้ำตาล")) {
          t.isSso = true;
          hasFbs = true;
        }
      });
      if (!hasFbs && programGroup !== "โปรแกรม MGR") {
        tests.push({ name: "ตรวจน้ำตาลในเลือด FBS (งดน้ำและอาหาร)", npo: true, isSso: true });
      }
    }
    
    // 4. Kidney Cr (age >= 35)
    if (age >= 35) {
      let hasKidney = false;
      tests.forEach(t => {
        if (programGroup === "โปรแกรม MGR") return;
        if (t.name.includes("Cr") || t.name.includes("ไต") || t.name.includes("BUN")) {
          t.isSso = true;
          hasKidney = true;
        }
      });
      if (!hasKidney && programGroup !== "โปรแกรม MGR") {
        tests.push({ name: "การทำงานของไต Cr และ eGFR", npo: false, isSso: true });
      }
    }
    
    // 5. Lipids
    if (age >= 20 && age < 35) {
      // 2 items: cholesterol and HDL
      tests.push({ name: "ตรวจไขมันในเลือด (Cholesterol, HDL) (งดน้ำและอาหาร)", npo: true, isSso: true });
    }
    
    // 6. HbsAg (age >= 35)
    if (age >= 35) {
      let hasHbs = false;
      tests.forEach(t => {
        if (t.name.includes("HbsAg") || t.name.includes("ตับอักเสบ")) {
          t.isSso = true;
          hasHbs = true;
        }
      });
      if (!hasHbs) {
        tests.push({ name: "เชื้อไวรัสตับอักเสบ HbsAg", npo: false, isSso: true });
      }
    }
  }
  
  // Render the combined tests list
  tests.forEach((t, i) => {
    const item = document.createElement("div");
    const isXray = t.name.includes("X-RAY") || t.name.includes("X-ray") || t.name.includes("เอกซเรย์");
    
    if (isXray && isPregnant) {
      item.className = "checkup-item pregnancy-disabled";
      item.innerHTML = `
        <i class="fa-solid fa-circle-xmark"></i>
        <span>${i + 1}. ${t.name} (งดตรวจเนื่องจากตั้งครรภ์)</span>
        <span class="pregnancy-badge">งดตรวจ</span>
      `;
    } else {
      let nameDisplay = t.name;
      if (t.isSso) {
        // Append asterisk
        nameDisplay = `${t.name} *`;
      }
      
      // Styling class
      if (t.isSso) {
        item.className = t.npo ? "checkup-item npo sso-merged-item" : "checkup-item sso-merged-item";
      } else {
        item.className = t.npo ? "checkup-item npo" : "checkup-item";
      }
      
      let iconClass = t.npo ? "fa-solid fa-triangle-exclamation" : "fa-solid fa-circle-check";
      let extraSpan = t.npo ? `<span class="npo-badge">งดน้ำ-งดอาหาร</span>` : "";
      
      item.innerHTML = `
        <i class="${iconClass}"></i>
        <span>${i + 1}. ${nameDisplay}</span>
        ${extraSpan}
      `;
    }
    itemsList.appendChild(item);
  });
  
  // Show / Hide the blue SSO note below checklist
  const ssoNote = document.getElementById("checkup-list-note");
  if (ssoNote) {
    ssoNote.style.display = useSso ? "block" : "none";
  }
  
  // Append Custom Risk Factor checkups
  const riskItems = employee.riskProgram ? employee.riskProgram.split(',').map(s => s.trim()).filter(Boolean) : [];
  if (riskItems.length > 0) {
    const startIdx = tests.length + 1;
    riskItems.forEach((riskText, idx) => {
      const item = document.createElement("div");
      item.className = "checkup-item risk-item";
      item.innerHTML = `
        <i class="fa-solid fa-stethoscope"></i>
        <span>${startIdx + idx}. ${riskText}</span>
        <span class="risk-badge">ปัจจัยเสี่ยง</span>
      `;
      itemsList.appendChild(item);
    });
  }
}



// --- Cancer Radio Option Card Selection ---
function selectCancerOption(labelElement) {
  document.querySelectorAll(".cancer-option").forEach(opt => {
    opt.classList.remove("selected");
  });
  
  labelElement.classList.add("selected");
  const input = labelElement.querySelector('input[type="radio"]');
  if (input) input.checked = true;
  validateFormCompletion();
}

// --- Phone Input validation (1-4 digits) ---
function validatePhone(input) {
  input.value = input.value.replace(/[^0-9]/g, "");
  validateFormCompletion();
}

// --- Handle search keypress (Enter key) ---
function handleLookupKeyPress(event) {
  if (event.key === "Enter") {
    event.preventDefault(); // Prevent form submission
    lookupEmployee();
  }
}

// --- Handle auto-search when 6 digits are typed ---
function handleLookupInput(input) {
  // Allow only numbers
  input.value = input.value.replace(/[^0-9]/g, "");
  
  if (input.value.length === 6) {
    input.blur(); // Dismiss keyboard
    lookupEmployee();
  }
}

// --- Shift or Location Change ---
function onShiftOrLocationChange() {
  const shift = document.getElementById("reg-shift").value;
  const location = document.getElementById("reg-location").value;
  
  const dateSelect = document.getElementById("reg-date");
  const gridContainer = document.getElementById("time-slots-grid-container");
  
  dateSelect.innerHTML = `<option value="" disabled selected>-- กรุณาเลือกสถานที่และกะก่อน --</option>`;
  dateSelect.disabled = true;
  gridContainer.innerHTML = `<div class="empty-slots-msg">กรุณาเลือกวันที่ตรวจสุขภาพก่อน</div>`;
  STATE.selectedTimeSlot = "";
  document.getElementById("selected-time-slot").value = "";
  
  if (!shift || !location) return;
  
  const matchedDates = STATE.configDates.filter(d => {
    if (d.location !== location) return false;
    
    if (shift === "ทีม A") {
      return d.team === "ทีม A";
    } else if (shift === "ทีม B") {
      return d.team === "ทีม B";
    } else if (shift === "เช้าตลอด" || shift === "คร่อมกะ") {
      return d.team === "ทีม A" || d.team === "ทีม B";
    }
    return false;
  });
  
  if (matchedDates.length === 0) {
    dateSelect.innerHTML = `<option value="" disabled selected>ไม่มีรอบวันเปิดตรวจสำหรับกะและสถานที่นี้</option>`;
    return;
  }
  
  dateSelect.innerHTML = `<option value="" disabled selected>-- เลือกวันที่ตรวจสุขภาพ --</option>`;
  const uniqueDates = [...new Set(matchedDates.map(md => md.dateString))];
  
  uniqueDates.forEach(dateStr => {
    const opt = document.createElement("option");
    opt.value = dateStr;
    opt.textContent = dateStr;
    dateSelect.appendChild(opt);
  });
  
  dateSelect.disabled = false;
  validateFormCompletion();
}

// --- Date Change: Render Time Slot Buttons Grid ---
function onDateChange() {
  renderTimeSlots(false); // Reset selection when date is manually changed
  validateFormCompletion();
}

function renderTimeSlots(preserveSelection = false) {
  const dateStr = document.getElementById("reg-date").value;
  const location = document.getElementById("reg-location").value;
  const gridContainer = document.getElementById("time-slots-grid-container");
  
  if (!dateStr || !location || !STATE.activeEmployee) {
    gridContainer.innerHTML = `<div class="empty-slots-msg">กรุณาเลือกสถานที่ กะทำงาน และวันที่ตรวจสุขภาพด้านบนก่อน เพื่อดึงข้อมูลรอบเวลาว่าง</div>`;
    if (!preserveSelection) {
      STATE.selectedTimeSlot = "";
      document.getElementById("selected-time-slot").value = "";
    }
    return;
  }
  
  const prevSelectedSlot = STATE.selectedTimeSlot;
  
  if (!preserveSelection) {
    STATE.selectedTimeSlot = "";
    document.getElementById("selected-time-slot").value = "";
  }
  
  gridContainer.innerHTML = "";
  
  STATE.configTimeSlots.forEach(slot => {
    const slotTime = slot.slotTime;
    const limit = slot.limit;
    
    const key = `${location}|${dateStr}|${slotTime}`;
    const booked = STATE.registrationCounts[key] || 0;
    const remaining = Math.max(0, limit - booked);
    
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "time-slot-btn";
    
    if (remaining <= 0) {
      // Disable slot since it's fully booked
      btn.disabled = true;
      btn.innerHTML = `
        <span class="time-text">${slotTime}</span>
        <span class="slot-status">เต็มแล้ว (จำกัด ${limit} คน)</span>
      `;
      // If the slot is now full, but was previously selected, clear selection and notify
      if (preserveSelection && prevSelectedSlot === slotTime) {
        STATE.selectedTimeSlot = "";
        document.getElementById("selected-time-slot").value = "";
        showToast(`รอบเวลา ${slotTime} เต็มแล้วค่ะ ระบบทำการยกเลิกตัวเลือกเดิมของท่าน`, "warning");
      }
    } else {
      // Available slot button
      btn.innerHTML = `
        <span class="time-text">${slotTime}</span>
        <span class="slot-status">เหลือ ${remaining}/${limit} ที่</span>
      `;
      btn.onclick = () => selectTimeSlot(slotTime, btn);
      
      // Restore selection if matching
      if (preserveSelection && prevSelectedSlot === slotTime) {
        btn.classList.add("selected");
      }
    }
    
    gridContainer.appendChild(btn);
  });
}

function selectTimeSlot(slotTime, buttonElement) {
  // Remove selected styling from all buttons in grid
  document.querySelectorAll(".time-slot-btn").forEach(btn => {
    btn.classList.remove("selected");
  });
  
  // Select clicked button
  buttonElement.classList.add("selected");
  
  // Update state & hidden input validation
  STATE.selectedTimeSlot = slotTime;
  document.getElementById("selected-time-slot").value = slotTime;
  validateFormCompletion();
}

// --- Submit Registration Flow ---
async function handleRegistrationSubmit(event) {
  event.preventDefault();
  
  if (!STATE.activeEmployee) {
    showToast("กรุณาค้นหารหัสพนักงานก่อน", "warning");
    return;
  }
  
  const phone = document.getElementById("reg-phone").value.trim();
  const shift = document.getElementById("reg-shift").value;
  const location = document.getElementById("reg-location").value;
  const dateStr = document.getElementById("reg-date").value;
  const timeStr = STATE.selectedTimeSlot;
  
  if (!phone || phone.length < 1 || phone.length > 4) {
    showToast("เบอร์โทรภายในต้องเป็นตัวเลข 1 - 4 หลัก", "warning");
    return;
  }
  
  if (!timeStr) {
    showToast("กรุณาเลือกรอบเวลาตรวจสุขภาพ", "warning");
    return;
  }
  
  let cancerTest = "";
  if (STATE.activeEmployee.programGroup === "โปรแกรม MGR") {
    const checkedRadio = document.querySelector('input[name="cancerTest"]:checked');
    if (!checkedRadio) {
      showToast("กรุณาเลือกรายการตรวจมะเร็ง 1 รายการ", "warning");
      return;
    }
    cancerTest = checkedRadio.value;
  }
  
  const ssoConsent = "ยินยอมใช้สิทธิ์ประกันสังคม";
  
  const isPregnant = !!STATE.activeEmployee.isPregnant;
  const payload = {
    employeeId: STATE.activeEmployee.employeeId,
    firstName: STATE.activeEmployee.firstName,
    lastName: STATE.activeEmployee.lastName,
    department: STATE.activeEmployee.department,
    phone: phone,
    shift: shift,
    location: location,
    dateString: dateStr,
    timeString: timeStr,
    cancerTest: cancerTest,
    riskProgram: STATE.activeEmployee.riskProgram || "",
    isPregnant: isPregnant,
    ssoConsent: ssoConsent
  };
  
  showLoader("กำลังบันทึกข้อมูลลงทะเบียนตรวจสุขภาพ...");
  
  if (CONFIG.currentMode === "mock") {
    await new Promise(resolve => setTimeout(resolve, 600));
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    
    // Find index to update or append
    const existingIndex = regs.findIndex(r => r.employeeId === payload.employeeId);
    
    // Validate slot limit locally
    const key = `${location}|${dateStr}|${timeStr}`;
    let isSameSlot = false;
    if (existingIndex !== -1) {
      const old = regs[existingIndex];
      if (old.location === location && old.dateString === dateStr && old.timeString === timeStr) {
        isSameSlot = true;
      }
    }
    
    const count = regs.filter(r => `${r.location}|${r.dateString}|${r.timeString}` === key).length;
    const activeSlot = STATE.configTimeSlots.find(s => s.slotTime === timeStr);
    const limit = activeSlot ? activeSlot.limit : 50;
    
    if (!isSameSlot && count >= limit) {
      hideLoader();
      showToast(`รอบเวลา ${timeStr} สำหรับ ${location} เต็มแล้ว!`, "error");
      return;
    }
    
    payload.timestamp = new Date().toLocaleString("th-TH");
    
    if (existingIndex !== -1) {
      regs[existingIndex] = payload;
    } else {
      regs.push(payload);
    }
    
    localStorage.setItem("MOCK_REGISTRATIONS", JSON.stringify(regs));
    await loadConfigAndCounts();
    hideLoader();
    showSuccessOverlay();
    resetForm();
  } else {
    let retries = 3;
    let saved = false;
    let attempt = 0;
    
    while (attempt < retries && !saved) {
      attempt++;
      try {
        if (attempt > 1) {
          showLoader(`ระบบหนาแน่นชั่วคราว กำลังพยายามอีกครั้ง (รอบที่ ${attempt}/${retries})...`);
          await new Promise(r => setTimeout(r, 2000)); // Wait 2 seconds before retry
        }
        
        const res = await callApi("saveRegistration", [payload]);
        if (res && res.success) {
          saved = true;
          await loadConfigAndCounts();
          hideLoader();
          showSuccessOverlay();
          resetForm();
        } else if (res && res.error && (res.error.includes("Lock timeout") || res.error.includes("หนาแน่น") || res.error.includes("limit exceeded"))) {
          console.warn(`Attempt ${attempt} failed with lock timeout: ${res.error}`);
          if (attempt >= retries) throw new Error(res.error);
        } else {
          throw new Error(res.error || "บันทึกข้อมูลล้มเหลว");
        }
      } catch (err) {
        if (attempt >= retries) {
          console.error(err);
          hideLoader();
          showToast(`เกิดข้อผิดพลาดในการลงทะเบียน: ${err.message}`, "error");
          break;
        }
      }
    }
  }
}

// --- Check Registration Status Tab Handler ---
async function checkRegistrationStatus() {
  const inputEl = document.getElementById("status-emp-id");
  let empId = inputEl.value.trim();
  
  if (!empId) {
    showToast("กรุณากรอกรหัสพนักงานที่ต้องการตรวจสอบ", "warning");
    return;
  }
  
  if (/^\d+$/.test(empId)) {
    empId = empId.padStart(6, '0');
    inputEl.value = empId;
  }
  
  showLoader("กำลังค้นหาข้อมูลการลงทะเบียน...");
  
  if (CONFIG.currentMode === "mock") {
    await new Promise(resolve => setTimeout(resolve, 500));
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    const foundReg = regs.find(r => r.employeeId === empId);
    if (foundReg) {
      const emp = MOCK_EMPLOYEES.find(e => e.employeeId === empId);
      if (emp) {
        foundReg.programGroup = emp.programGroup;
        foundReg.age = emp.age;
        foundReg.gender = emp.gender || "";
      }
    }
    
    hideLoader();
    renderStatusCard(foundReg, empId);
  } else {
    try {
      const res = await callApi("getRegistrationByEmpId", [empId]);
      if (res && res.success) {
        hideLoader();
        renderStatusCard(res.data, empId);
      } else {
        throw new Error(res.error || "ดึงข้อมูลล้มเหลว");
      }
    } catch (err) {
      console.error(err);
      hideLoader();
      showToast(`มีข้อผิดพลาด: ${err.message}`, "error");
      renderStatusCard(null, empId);
    }
  }
}

function handleStatusSearchKeyPress(event) {
  if (event.key === "Enter") {
    checkRegistrationStatus();
  }
}

function renderStatusCard(reg, searchId) {
  const cardContainer = document.getElementById("result-card-container");
  
  if (!reg) {
    showToast(`ไม่พบข้อมูลการลงทะเบียนสำหรับรหัสพนักงาน ${searchId} ในระบบ`, "error");
    cardContainer.classList.remove("visible");
    STATE.activeRegistration = null;
    return;
  }
  
  // Store globally to handle Edit/Cancel actions
  STATE.activeRegistration = reg;
  
  // Populate Card
  document.getElementById("card-empid-val").textContent = reg.employeeId;
  document.getElementById("card-name-val").textContent = `${reg.firstName} ${reg.lastName}`;
  document.getElementById("card-dept-phone-val").textContent = `${reg.department} / เบอร์ภายใน ${reg.phone}`;
  
  document.getElementById("card-loc-val").textContent = reg.location === "LPN1" ? "LPN1 (Lobby(เก่า) อาคาร2)" : "LPN2 (ห้องฝึกอบรมชั้น 2)";
  document.getElementById("card-date-val").textContent = reg.dateString;
  document.getElementById("card-time-val").textContent = `${reg.timeString} น.`;
  
  document.getElementById("card-program-title").textContent = reg.programGroup;
  
  // Render test items checklist exactly like main list
  const checklistContainer = document.getElementById("card-tests-list-container");
  checklistContainer.innerHTML = "";
  
  const useSso = true; // Always display SSO items on ticket
  
  const programGroup = reg.programGroup;
  const baseTests = PROGRAM_TESTS[programGroup] || [];
  
  // Clone base tests
  let tests = baseTests.map(t => ({ ...t, isSso: false }));
  
  const age = parseInt(reg.age || 0, 10);
  const isFemale = reg.gender === "F" || reg.gender === "Female" || !!reg.isPregnant;
  
  if (useSso) {
    // 3. FBS (age >= 35)
    if (age >= 35) {
      let hasFbs = false;
      tests.forEach(t => {
        if (programGroup === "โปรแกรม MGR") return;
        if (t.name.includes("FBS") || t.name.includes("น้ำตาล")) {
          t.isSso = true;
          hasFbs = true;
        }
      });
      if (!hasFbs && programGroup !== "โปรแกรม MGR") {
        tests.push({ name: "ตรวจน้ำตาลในเลือด FBS (งดน้ำและอาหาร)", npo: true, isSso: true });
      }
    }
    
    // 4. Kidney Cr (age >= 35)
    if (age >= 35) {
      let hasKidney = false;
      tests.forEach(t => {
        if (programGroup === "โปรแกรม MGR") return;
        if (t.name.includes("Cr") || t.name.includes("ไต") || t.name.includes("BUN")) {
          t.isSso = true;
          hasKidney = true;
        }
      });
      if (!hasKidney && programGroup !== "โปรแกรม MGR") {
        tests.push({ name: "การทำงานของไต Cr และ eGFR", npo: false, isSso: true });
      }
    }
    
    // 5. Lipids
    if (age >= 20 && age < 35) {
      tests.push({ name: "ตรวจไขมันในเลือด (Cholesterol, HDL) (งดน้ำและอาหาร)", npo: true, isSso: true });
    }
    
    // 6. HbsAg (age >= 35)
    if (age >= 35) {
      let hasHbs = false;
      tests.forEach(t => {
        if (t.name.includes("HbsAg") || t.name.includes("ตับอักเสบ")) {
          t.isSso = true;
          hasHbs = true;
        }
      });
      if (!hasHbs) {
        tests.push({ name: "เชื้อไวรัสตับอักเสบ HbsAg", npo: false, isSso: true });
      }
    }
  }
  
  tests.forEach((t, index) => {
    const item = document.createElement("div");
    const isXray = t.name.includes("X-RAY") || t.name.includes("X-ray") || t.name.includes("เอกซเรย์");
    
    if (isXray && reg.isPregnant) {
      item.className = "ticket-test-item ticket-pregnancy-disabled";
      item.innerHTML = `
        <i class="fa-solid fa-circle-xmark"></i>
        <span>${index + 1}. ${t.name} *งดตรวจเนื่องจากตั้งครรภ์*</span>
      `;
    } else {
      let nameDisplay = t.name;
      if (t.isSso) {
        nameDisplay = `${t.name} *`;
        item.className = "ticket-test-item sso-merged-item";
      } else {
        item.className = "ticket-test-item";
      }
      
      // Add extra text for NPO fasting items
      let testNameDisplay = `${index + 1}. ${nameDisplay}`;
      if (t.npo) {
        if (!testNameDisplay.includes("งดน้ำ") && !testNameDisplay.includes("งดอาหาร")) {
          testNameDisplay += " *งดน้ำงดอาหาร*";
        }
      }
      
      item.innerHTML = `
        <i class="fa-solid fa-check"></i>
        <span>${testNameDisplay}</span>
      `;
    }
    checklistContainer.appendChild(item);
  });
  
  // Append Custom Risk Factor checkups on Ticket
  const riskItems = reg.riskProgram ? reg.riskProgram.split(',').map(s => s.trim()).filter(Boolean) : [];
  let currentIdx = tests.length + 1;
  riskItems.forEach((riskText) => {
    const item = document.createElement("div");
    item.className = "ticket-test-item ticket-risk-item";
    item.innerHTML = `
      <i class="fa-solid fa-stethoscope"></i>
      <span>${currentIdx}. ${riskText} *ปัจจัยเสี่ยง*</span>
    `;
    checklistContainer.appendChild(item);
    currentIdx++;
  });
  
  // Append Selected Cancer Checkup if manager
  if (reg.programGroup === "โปรแกรม MGR" && reg.cancerTest) {
    const item = document.createElement("div");
    item.className = "ticket-test-item cancer-gold";
    item.innerHTML = `
      <i class="fa-solid fa-crown"></i>
      <span>${currentIdx}. ${reg.cancerTest}</span>
    `;
    checklistContainer.appendChild(item);
    currentIdx++;
  }
  
  const cardSsoSection = document.getElementById("card-sso-section");
  if (cardSsoSection) {
    cardSsoSection.style.display = "none";
  }
  
  cardContainer.classList.add("visible");
  cardContainer.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// --- Edit Booking Action ---
function editRegistration() {
  if (!STATE.activeRegistration) return;
  
  const reg = STATE.activeRegistration;
  
  // Switch to Register Tab
  switchTab("register");
  
  // Trigger Employee Lookup
  const empInput = document.getElementById("reg-emp-id");
  empInput.value = reg.employeeId;
  
  lookupEmployee().then(() => {
    // Wait for employee details to load and fields to render, then populate choices
    setTimeout(() => {
      if (STATE.activeEmployee) {
        document.getElementById("reg-phone").value = reg.phone;
        document.getElementById("reg-shift").value = reg.shift;
        document.getElementById("reg-location").value = reg.location;
        
        // Trigger shift change calculation
        onShiftOrLocationChange();
        
        // Populate date choice
        document.getElementById("reg-date").value = reg.dateString;
        onDateChange();
        
        // Find and select the time slot button
        setTimeout(() => {
          const buttons = document.querySelectorAll(".time-slot-btn");
          buttons.forEach(btn => {
            const timeText = btn.querySelector(".time-text").textContent;
            if (timeText === reg.timeString) {
              selectTimeSlot(reg.timeString, btn);
            }
          });
        }, 100);
        
        // Pre-fill cancer selection if manager
        if (reg.programGroup === "โปรแกรม MGR" && reg.cancerTest) {
          const radio = document.querySelector(`input[name="cancerTest"][value="${reg.cancerTest}"]`);
          if (radio) {
            radio.checked = true;
            // Highlight wrapper
            const parentLabel = radio.closest(".cancer-option");
            if (parentLabel) selectCancerOption(parentLabel);
          }
        }
        
        // Render checkup list according to saved pregnancy status
        renderCheckupList(STATE.activeEmployee, !!reg.isPregnant);
        
        showToast("โหลดข้อมูลเดิมให้คุณแก้ไขแล้วค่ะ", "success");
      }
    }, 600);
  });
}

// --- Cancel Booking Action ---
async function cancelRegistration() {
  if (!STATE.activeRegistration) return;
  
  const reg = STATE.activeRegistration;
  const confirmCancel = confirm(`คุณต้องการยกเลิกการลงทะเบียนตรวจสุขภาพสำหรับรหัสพนักงาน ${reg.employeeId} หรือไม่?`);
  
  if (!confirmCancel) return;
  
  showLoader("กำลังยกเลิกการลงทะเบียนของคุณ...");
  
  if (CONFIG.currentMode === "mock") {
    await new Promise(resolve => setTimeout(resolve, 600));
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    
    // Filter out registration
    const updatedRegs = regs.filter(r => r.employeeId !== reg.employeeId);
    localStorage.setItem("MOCK_REGISTRATIONS", JSON.stringify(updatedRegs));
    
    await loadConfigAndCounts();
    hideLoader();
    showToast("ยกเลิกการลงทะเบียนสำเร็จแล้วค่ะ", "success");
    
    // Hide ticket card
    document.getElementById("result-card-container").classList.remove("visible");
    STATE.activeRegistration = null;
    document.getElementById("status-emp-id").value = "";
  } else {
    let retries = 3;
    let deleted = false;
    let attempt = 0;
    
    while (attempt < retries && !deleted) {
      attempt++;
      try {
        if (attempt > 1) {
          showLoader(`ระบบหนาแน่นชั่วคราว กำลังพยายามอีกครั้ง (รอบที่ ${attempt}/${retries})...`);
          await new Promise(r => setTimeout(r, 2000)); // Wait 2 seconds before retry
        }
        
        const res = await callApi("deleteRegistration", [reg.employeeId]);
        if (res && res.success) {
          deleted = true;
          await loadConfigAndCounts();
          hideLoader();
          showToast("ยกเลิกการลงทะเบียนสำเร็จแล้วค่ะ", "success");
          
          document.getElementById("result-card-container").classList.remove("visible");
          STATE.activeRegistration = null;
          document.getElementById("status-emp-id").value = "";
        } else if (res && res.error && (res.error.includes("Lock timeout") || res.error.includes("หนาแน่น") || res.error.includes("limit exceeded"))) {
          console.warn(`Attempt ${attempt} failed with lock timeout: ${res.error}`);
          if (attempt >= retries) throw new Error(res.error);
        } else {
          throw new Error(res.error || "ไม่สามารถยกเลิกได้");
        }
      } catch (err) {
        if (attempt >= retries) {
          console.error(err);
          hideLoader();
          showToast(`มีข้อผิดพลาดในการยกเลิก: ${err.message}`, "error");
          break;
        }
      }
    }
  }
}

// --- Helper Functions: Reset, Alerts, API Client ---
function resetForm() {
  document.getElementById("health-registration-form").reset();
  document.getElementById("employee-profile-box").style.display = "none";
  document.getElementById("selected-time-slot").value = "";
  const prevRegCard = document.getElementById("prev-reg-card");
  if (prevRegCard) prevRegCard.style.display = "none";
  STATE.activeEmployee = null;
  STATE.selectedTimeSlot = "";
  validateFormCompletion();
}

function validateFormCompletion() {
  const submitBtn = document.getElementById("btn-submit-registration");
  if (!submitBtn) return;
  
  if (!STATE.activeEmployee) {
    submitBtn.disabled = true;
    return;
  }
  
  // Re-render checklist (SSO items are always merged now)
  renderCheckupList(STATE.activeEmployee, !!STATE.activeEmployee.isPregnant);
  
  const phone = document.getElementById("reg-phone").value.trim();
  const isPhoneValid = /^\d{1,4}$/.test(phone);
  
  const shift = document.getElementById("reg-shift").value;
  const location = document.getElementById("reg-location").value;
  const dateStr = document.getElementById("reg-date").value;
  const timeStr = STATE.selectedTimeSlot;
  
  let isCancerValid = true;
  if (STATE.activeEmployee.programGroup === "โปรแกรม MGR") {
    isCancerValid = document.querySelector('input[name="cancerTest"]:checked') !== null;
  }
  
  const isFormComplete = isPhoneValid && shift && location && dateStr && timeStr && isCancerValid;
  submitBtn.disabled = !isFormComplete;
}

function showToast(text, type = "info") {
  const toast = document.getElementById("toast-msg");
  const icon = document.getElementById("toast-icon");
  const textEl = document.getElementById("toast-text");
  
  toast.className = "toast-msg";
  icon.className = "fa-solid";
  
  if (type === "success") {
    toast.classList.add("success");
    icon.classList.add("fa-circle-check");
  } else if (type === "error") {
    toast.classList.add("error");
    icon.classList.add("fa-circle-xmark");
  } else if (type === "warning") {
    toast.classList.add("warning");
    icon.classList.add("fa-triangle-exclamation");
  } else {
    toast.classList.add("info");
    icon.classList.add("fa-circle-info");
  }
  
  textEl.textContent = text;
  toast.classList.add("show");
  
  setTimeout(() => {
    toast.classList.remove("show");
  }, 4000);
}

function showLoader(message = "กำลังประมวลผล...") {
  document.getElementById("loading-text").textContent = message;
  document.getElementById("loading-overlay").classList.add("active");
}

function hideLoader() {
  document.getElementById("loading-overlay").classList.remove("active");
}

function showSuccessOverlay() {
  document.getElementById("success-overlay").classList.add("active");
}

function closeSuccessOverlay() {
  document.getElementById("success-overlay").classList.remove("active");
}

async function callApi(action, args) {
  if (!CONFIG.apiUrl || CONFIG.apiUrl.includes("your_deployed_url")) {
    showToast("กรุณาติดตั้ง Web App URL ของ Google Sheets ในไฟล์ app.js ก่อน", "error");
    return { success: false, error: "API URL not configured" };
  }
  
  try {
    const payload = {
      action: action,
      args: args
    };
    
    const response = await fetch(CONFIG.apiUrl, {
      method: "POST",
      mode: "cors",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(payload)
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    return result;
  } catch (err) {
    console.error("API Call error:", err);
    throw err;
  }
}

// --- Real-time Synchronisation Manager ---
const SYNC_CONFIG = {
  minIntervalMs: 6000,
  maxIntervalMs: 9000,
  timerId: null,
  isSyncing: false
};

function initRealTimeSync() {
  if (SYNC_CONFIG.timerId) clearTimeout(SYNC_CONFIG.timerId);
  
  // Start the dynamic scheduling loop
  scheduleNextSync();
}

function scheduleNextSync() {
  // Generate random jittered interval (e.g. between 6s and 9s) to distribute server load
  const jitteredInterval = Math.floor(Math.random() * (SYNC_CONFIG.maxIntervalMs - SYNC_CONFIG.minIntervalMs + 1)) + SYNC_CONFIG.minIntervalMs;
  
  SYNC_CONFIG.timerId = setTimeout(async () => {
    // Only run sync if page is visible and user is actively looking at slot options
    if (!document.hidden && shouldSyncCounts()) {
      SYNC_CONFIG.isSyncing = true;
      try {
        await performBackgroundSync();
      } catch (err) {
        console.warn("Background sync failed:", err);
      } finally {
        SYNC_CONFIG.isSyncing = false;
      }
    }
    // Loop
    scheduleNextSync();
  }, jitteredInterval);
}

function shouldSyncCounts() {
  // Only poll if on registration tab and date/location are selected (user looking at slots grid)
  const regTab = document.getElementById("tab-register");
  const isRegTabActive = regTab && regTab.classList.contains("active");
  if (!isRegTabActive) return false;
  
  const dateStr = document.getElementById("reg-date").value;
  const location = document.getElementById("reg-location").value;
  return !!(dateStr && location && STATE.activeEmployee);
}

async function performBackgroundSync() {
  if (CONFIG.currentMode === "mock") {
    // Read from local mock registry in LocalStorage
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    STATE.registrationCounts = {};
    regs.forEach(r => {
      const key = `${r.location}|${r.dateString}|${r.timeString}`;
      STATE.registrationCounts[key] = (STATE.registrationCounts[key] || 0) + 1;
    });
    // Trigger in-place update of slots grid
    renderTimeSlots(true);
    pulseRealtimeIndicator();
  } else {
    // Call Google Sheets API silently (without displaying full page blocker loader)
    const response = await callApi("getConfigAndSlots", []);
    if (response && response.success) {
      STATE.configDates = response.data.dates;
      STATE.configTimeSlots = response.data.timeSlots;
      STATE.registrationCounts = response.data.registrationCounts;
      // Trigger in-place update of slots grid
      renderTimeSlots(true);
      pulseRealtimeIndicator();
    }
  }
}

function pulseRealtimeIndicator() {
  const badge = document.getElementById("realtime-sync-status");
  if (!badge) return;
  
  badge.classList.add("syncing");
  const text = badge.querySelector(".sync-text");
  const dot = badge.querySelector(".pulse-dot");
  
  if (text) text.textContent = "อัปเดตล่าสุด: เมื่อครู่";
  if (dot) dot.style.backgroundColor = "var(--primary-dark)";
  
  setTimeout(() => {
    badge.classList.remove("syncing");
    if (text) text.textContent = "เชื่อมต่อเรียลไทม์";
    if (dot) dot.style.backgroundColor = "var(--success-color)";
  }, 1500);
}

// --- Special Self-Pay Catalog Page Logic ---
function renderSpecialCatalog() {
  const grid = document.getElementById("catalog-cards-grid");
  if (!grid) return;
  
  grid.innerHTML = "";
  
  // Filter the items based on Category & Search text
  const query = STATE.catalogSearchQuery.toLowerCase().trim();
  const category = STATE.catalogCategory;
  
  // Determine if the currently loaded activeEmployee is eligible for age 50+ welfare
  const employeeAge = STATE.activeEmployee ? STATE.activeEmployee.age : 0;
  
  const filtered = SPECIAL_TESTS.filter(item => {
    // 1. Search Query filtering
    if (query) {
      const matchName = item.name.toLowerCase().includes(query);
      const matchPurpose = item.purpose.toLowerCase().includes(query);
      const matchCategory = item.category.toLowerCase().includes(query);
      if (!matchName && !matchPurpose && !matchCategory) return false;
    }
    
    // 2. Category filtering
    if (category === "all") return true;
    if (category === "welfare") return !!item.welfare;
    return item.category === category;
  });
  
  if (filtered.length === 0) {
    grid.innerHTML = `<div class="empty-slots-msg" style="grid-column: 1 / -1;">ไม่พบรายการตรวจสุขภาพพิเศษตามเงื่อนไขการค้นหา</div>`;
    return;
  }
  
  filtered.forEach(item => {
    const card = document.createElement("div");
    card.className = "catalog-card";
    
    // Fasting label
    let fastingBadge = "";
    if (item.fasting) {
      fastingBadge = `<span class="catalog-badge fasting"><i class="fa-solid fa-cookie-bite"></i> งดน้ำ-อาหาร 8-12 ชม.</span>`;
    }
    
    // Location label if custom
    let locBadge = "";
    if (item.location) {
      locBadge = `<span class="catalog-badge location"><i class="fa-solid fa-location-dot"></i> ${item.location}</span>`;
    }
    
    // Welfare label and welfare price calculation box
    let welfareBadge = "";
    let welfareBox = "";
    if (item.welfare) {
      welfareBadge = `<span class="catalog-badge welfare"><i class="fa-solid fa-gift"></i> สวัสดิการช่วยเหลือ 50%</span>`;
      
      // Calculate dynamic price based on whether activeEmployee is logged in and meets the criteria
      let isEligible = false;
      let reason = "จำกัดเฉพาะพนักงานอายุ 50 ปีขึ้นไป";
      
      if (STATE.activeEmployee) {
        const meetsAge = STATE.activeEmployee.age >= item.welfare.minAge;
        let meetsLevel = true;
        
        if (item.welfare.minLevel) {
          // If a level check is required (e.g. M4/T5 or M5/T6), check programName or programGroup
          const userProg = STATE.activeEmployee.programName || "";
          if (item.welfare.minLevel === "M5/T6") {
            // M5/T6 corresponds to Program MGR
            meetsLevel = (userProg.includes("MGR") || userProg.toLowerCase().includes("mgr"));
          } else if (item.welfare.minLevel === "M4/T5") {
            // M4/T5 corresponds to MGR or potentially senior levels.
            meetsLevel = (userProg.includes("MGR") || userProg.toLowerCase().includes("mgr") || userProg.includes("35 ปีขึ้นไป") || STATE.activeEmployee.age >= 35);
          }
        }
        
        if (!meetsAge) {
          reason = `ไม่ตรงตามเกณฑ์สิทธิ์สวัสดิการ (สำหรับอายุ ${item.welfare.minAge} ปีขึ้นไป)`;
        } else if (!meetsLevel) {
          reason = `พนักงานระดับตำแหน่งไม่ถึงเกณฑ์สิทธิ์ (${item.welfare.minLevel})`;
        } else {
          isEligible = true;
        }
      } else {
        reason = "กรุณากรอกรหัสพนักงานในแท็บลงทะเบียนเพื่อตรวจสอบสิทธิ์สวัสดิการ";
      }
      
      // Net calculations
      // 50% co-pay, but co-pay amount cannot exceed the welfare limit
      let basePrice = item.price;
      let copayText = "";
      
      if (item.welfare.isMri) {
        // Special case for MRI Brain (M5/T6, 50% max 10,000)
        const companyPart = 10000; 
        const empPartBrain = 24000 - companyPart;
        const empPartBoth = 29000 - companyPart;
        
        if (isEligible) {
          copayText = `
            <div><span>MRI Brain (24,000 ฿):</span><span class="text-green">จ่ายเพียง ${formatPrice(empPartBrain)} ฿</span></div>
            <div><span>MRI+MRA (29,000 ฿):</span><span class="text-green">จ่ายเพียง ${formatPrice(empPartBoth)} ฿</span></div>
            <div style="font-size: 0.75rem; color: #64748b; margin-top: 4px; border-top: 1px dashed #fcd34d; padding-top: 4px;">* บริษัทช่วยเหลือสูงสุด ${formatPrice(companyPart)} ฿</div>
          `;
        } else {
          copayText = `
            <div style="color: #b91c1c; font-weight: 700;"><i class="fa-solid fa-circle-xmark"></i> ไม่มีสิทธิ์ใช้สวัสดิการ:</div>
            <div style="font-size: 0.78rem; color: #7f1d1d; margin-top: 2px;">${reason}</div>
          `;
        }
      } else if (item.welfare.isColon) {
        // Special case for Colonoscopy (M5/T6, 50% max 15,000)
        const companyPart = Math.min(basePrice * 0.5, item.welfare.limit);
        const empPart = basePrice - companyPart;
        const companyPartBiopsy = Math.min(15000 * 0.5, item.welfare.limit);
        const empPartBiopsy = 15000 - companyPartBiopsy;
        
        if (isEligible) {
          copayText = `
            <div><span>ส่องกล้องปกติ (14,000 ฿):</span><span class="text-green">จ่ายเพียง ${formatPrice(empPart)} ฿</span></div>
            <div><span>รวมตัดชิ้นเนื้อ (15,000 ฿):</span><span class="text-green">จ่ายเพียง ${formatPrice(empPartBiopsy)} ฿</span></div>
            <div style="font-size: 0.75rem; color: #64748b; margin-top: 4px; border-top: 1px dashed #fcd34d; padding-top: 4px;">* บริษัทสมทบคนละครึ่ง 50% สูงสุด 15,000 ฿</div>
          `;
        } else {
          copayText = `
            <div style="color: #b91c1c; font-weight: 700;"><i class="fa-solid fa-circle-xmark"></i> ไม่มีสิทธิ์ใช้สวัสดิการ:</div>
            <div style="font-size: 0.78rem; color: #7f1d1d; margin-top: 2px;">${reason}</div>
          `;
        }
      } else {
        // Normal co-pay calculations
        const companyPart = Math.min(basePrice * 0.5, item.welfare.limit);
        const empPart = basePrice - companyPart;
        
        if (isEligible) {
          copayText = `
            <div><span>ราคาตรวจปกติ:</span><span>${formatPrice(basePrice)} ฿</span></div>
            <div><span>บริษัทสมทบช่วยเหลือ (50%):</span><span>-${formatPrice(companyPart)} ฿</span></div>
            <div class="welfare-net-row"><span>ราคาพนักงานจ่ายจริง:</span><span>${formatPrice(empPart)} ฿</span></div>
          `;
        } else {
          copayText = `
            <div style="color: #b91c1c; font-weight: 700;"><i class="fa-solid fa-circle-xmark"></i> ไม่มีสิทธิ์ใช้สวัสดิการ:</div>
            <div style="font-size: 0.78rem; color: #7f1d1d; margin-top: 2px;">${reason}</div>
          `;
        }
      }
      
      welfareBox = `
        <div class="catalog-welfare-box">
          ${copayText}
        </div>
      `;
    } else if (item.id === 2) {
      // Special note for cholesterol (lipid) - free for 35+
      let isFree = false;
      if (STATE.activeEmployee) {
        isFree = (STATE.activeEmployee.age >= 35);
      }
      
      welfareBox = `
        <div class="catalog-welfare-box" style="background: #ecfdf5; border-color: #a7f3d0; color: #065f46;">
          <div style="font-weight: 800; display: flex; align-items: center; gap: 4px;"><i class="fa-solid fa-circle-check"></i> สิทธิ์พนักงานอายุ 35 ปีขึ้นไป:</div>
          <div style="margin-top: 4px; display: flex; justify-content: space-between; font-size: 0.82rem;">
            <span>สิทธิ์ฟรีสวัสดิการ:</span>
            <span style="font-weight: bold; color: #047857;">${isFree ? "ได้รับสิทธิ์ตรวจฟรี 🎉" : "ไม่อยู่ในเกณฑ์ตรวจฟรี (ชำระเอง 160 ฿)"}</span>
          </div>
        </div>
      `;
    }
    
    // Render price value
    let priceDisplay = `${formatPrice(item.price)} ฿`;
    if (item.id === 33) {
      priceDisplay = "เริ่มต้น 24,000 ฿";
    }
    
    card.innerHTML = `
      <div class="catalog-card-top">
        <span class="catalog-card-id">รายการลำดับที่ ${item.id}</span>
        <h4 class="catalog-card-name">${item.name}</h4>
        <p class="catalog-card-purpose">${item.purpose}</p>
        
        <div class="catalog-badges">
          ${fastingBadge}
          ${locBadge}
          ${welfareBadge}
        </div>
      </div>
      
      <div class="catalog-card-bottom">
        <div class="catalog-price-row">
          <span class="price-label">ราคาตรวจปกติ</span>
          <span class="price-value">${priceDisplay}</span>
        </div>
        ${welfareBox}
      </div>
    `;
    
    grid.appendChild(card);
  });
}

function formatPrice(number) {
  return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function filterCatalog() {
  const searchInput = document.getElementById("catalog-search-input");
  if (searchInput) {
    STATE.catalogSearchQuery = searchInput.value;
    renderSpecialCatalog();
  }
}

function setCatalogCategory(catName) {
  STATE.catalogCategory = catName;
  
  // Highlight active filter button
  const container = document.getElementById("catalog-filter-pills");
  if (container) {
    container.querySelectorAll(".filter-pill").forEach(btn => {
      btn.classList.remove("active");
    });
    
    // Find matching button based on click category name
    const buttons = container.querySelectorAll(".filter-pill");
    const catMap = {
      'all': 'ทั้งหมด',
      'blood': 'ตรวจเลือด/ปัสสาวะ',
      'cancer': 'ตรวจมะเร็ง',
      'scan': 'สแกน/เฉพาะทาง',
      'welfare': 'มีสวัสดิการสมทบ'
    };
    
    buttons.forEach(btn => {
      if (btn.textContent.includes(catMap[catName])) {
        btn.classList.add("active");
      }
    });
  }
  
  renderSpecialCatalog();
}

// --- Admin Dashboard Logic ---

function checkAdminState() {
  const loginCard = document.getElementById("admin-login-card");
  const dashContent = document.getElementById("admin-dashboard-content");
  
  if (STATE.isAdminAuthenticated) {
    loginCard.style.display = "none";
    dashContent.style.display = "block";
    loadAdminDashboardData();
  } else {
    loginCard.style.display = "block";
    dashContent.style.display = "none";
    document.getElementById("admin-pass-input").value = "";
    document.getElementById("admin-pass-input").focus();
  }
}

function handleAdminLogin(event) {
  event.preventDefault();
  const passwordInput = document.getElementById("admin-pass-input");
  const password = passwordInput.value.trim();
  
  if (password === "ad2026") {
    STATE.isAdminAuthenticated = true;
    showToast("เข้าสู่ระบบแอดมินสำเร็จ", "success");
    checkAdminState();
  } else {
    showToast("รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง", "error");
    passwordInput.value = "";
    passwordInput.focus();
  }
}

function handleAdminLogout() {
  STATE.isAdminAuthenticated = false;
  STATE.adminDashboardData = null;
  showToast("ออกจากระบบแอดมินเรียบร้อยแล้ว", "info");
  checkAdminState();
}

async function loadAdminDashboardData() {
  showLoader("กำลังดึงข้อมูลแดชบอร์ด...");
  try {
    if (CONFIG.currentMode === "mock") {
      await new Promise(resolve => setTimeout(resolve, 600));
      const eligibleEmployees = MOCK_EMPLOYEES.filter(emp => {
        return !(emp.checkupRight && emp.checkupRight.indexOf("ไม่มีสิทธิ์") !== -1);
      });
      const registrations = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
      
      STATE.adminDashboardData = {
        employees: eligibleEmployees,
        registrations: registrations
      };
    } else {
      const res = await callApi("getAdminDashboardData", []);
      if (res && res.success) {
        STATE.adminDashboardData = res.data;
      } else {
        throw new Error(res.error || "ไม่สามารถดึงข้อมูลแดชบอร์ดได้");
      }
    }
    renderAdminDashboard();
  } catch (err) {
    console.error(err);
    showToast(`ดึงข้อมูลแดชบอร์ดล้มเหลว: ${err.message}`, "error");
  } finally {
    hideLoader();
  }
}

function renderAdminDashboard() {
  if (!STATE.adminDashboardData) return;
  
  const { employees, registrations } = STATE.adminDashboardData;
  
  const total = employees.length;
  const registeredList = registrations.filter(r => employees.some(e => e.employeeId === r.employeeId));
  const registeredCount = registeredList.length;
  const unregisteredCount = total - registeredCount;
  const percentage = total > 0 ? ((registeredCount / total) * 100).toFixed(1) : "0.0";
  
  document.getElementById("stat-total-emp").textContent = total;
  document.getElementById("stat-registered-emp").textContent = registeredCount;
  document.getElementById("stat-unregistered-emp").textContent = unregisteredCount;
  document.getElementById("stat-percent-emp").textContent = `${percentage}%`;
  
  // Populate dates select filter for slots table
  const dateFilterSelect = document.getElementById("admin-slots-date-filter");
  const uniqueDates = [...new Set(STATE.configDates.map(d => d.dateString))];
  let dateFilterHtml = '<option value="all">ทั้งหมดทุกวัน</option>';
  uniqueDates.forEach(dateStr => {
    dateFilterHtml += `<option value="${dateStr}">${dateStr}</option>`;
  });
  dateFilterSelect.innerHTML = dateFilterHtml;
  
  // Populate departments select filter for unregistered list
  const deptSelect = document.getElementById("admin-unreg-dept-select");
  const uniqueDepts = [...new Set(employees.map(e => e.department).filter(Boolean))].sort();
  let deptHtml = '<option value="all">ทั้งหมดทุกแผนก</option>';
  uniqueDepts.forEach(dept => {
    deptHtml += `<option value="${dept}">${dept}</option>`;
  });
  deptSelect.innerHTML = deptHtml;
  
  // Render sub-sections
  renderAdminSlotsDashboard();
  renderAdminUnregisteredList();
}

function renderAdminSlotsDashboard() {
  if (!STATE.adminDashboardData) return;
  
  const { registrations } = STATE.adminDashboardData;
  const locationFilter = document.getElementById("admin-slots-location-filter").value;
  const dateFilter = document.getElementById("admin-slots-date-filter").value;
  const tableBody = document.getElementById("admin-slots-table-body");
  
  tableBody.innerHTML = "";
  
  STATE.configDates.forEach(date => {
    // Apply filters
    if (locationFilter !== "all" && date.location !== locationFilter) return;
    if (dateFilter !== "all" && date.dateString !== dateFilter) return;
    
    // For each date, loop through each config time slot
    STATE.configTimeSlots.forEach(slot => {
      const slotTime = slot.slotTime;
      const limit = parseInt(slot.limit || 50, 10);
      
      // Count registrations matching location, date, time
      const count = registrations.filter(r => 
        r.location === date.location && 
        r.dateString === date.dateString && 
        r.timeString === slotTime
      ).length;
      
      const isFull = count >= limit;
      const statusBadgeHtml = isFull 
        ? '<span class="badge-status full">เต็ม</span>' 
        : `<span class="badge-status available">ว่าง (${limit - count})</span>`;
      
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong>${date.location}</strong></td>
        <td>${date.dateString}</td>
        <td>${date.team}</td>
        <td>${slotTime}</td>
        <td><strong>${count}</strong></td>
        <td>${limit}</td>
        <td>${statusBadgeHtml}</td>
      `;
      tableBody.appendChild(tr);
    });
  });
  
  if (tableBody.innerHTML === "") {
    tableBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: var(--text-muted); padding: 15px;">ไม่มีข้อมูลรอบตรวจที่ตรงกับตัวกรอง</td></tr>`;
  }
}

function renderAdminUnregisteredList() {
  if (!STATE.adminDashboardData) return;
  
  const { employees, registrations } = STATE.adminDashboardData;
  const deptFilter = document.getElementById("admin-unreg-dept-select").value;
  const tableBody = document.getElementById("admin-unreg-table-body");
  
  tableBody.innerHTML = "";
  
  // Find unregistered
  const unregisteredEmployees = employees.filter(emp => 
    !registrations.some(r => r.employeeId === emp.employeeId)
  );
  
  // Apply department filter
  const filtered = unregisteredEmployees.filter(emp => 
    deptFilter === "all" || emp.department === deptFilter
  );
  
  filtered.forEach(emp => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${emp.employeeId}</strong></td>
      <td>${emp.firstName} ${emp.lastName}</td>
      <td><span class="profile-dept-badge">${emp.department}</span></td>
      <td><span class="badge-status unregistered">ยังไม่ได้ลงทะเบียน</span></td>
      <td style="text-align: center;">
        <button type="button" class="btn-copy-individual" onclick="copyReminder('${emp.employeeId}')">
          <i class="fa-solid fa-copy"></i> คัดลอกคำเตือน
        </button>
      </td>
    `;
    tableBody.appendChild(tr);
  });
  
  if (filtered.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 20px;">ไม่มีรายชื่อพนักงานที่ค้างลงทะเบียนตามแผนกที่เลือก</td></tr>`;
  }
}

function switchAdminSubTab(subTabName) {
  STATE.adminSubTab = subTabName;
  
  document.querySelectorAll(".admin-sub-btn").forEach(btn => btn.classList.remove("active"));
  document.getElementById(`admin-sub-btn-${subTabName}`).classList.add("active");
  
  document.querySelectorAll(".admin-sec").forEach(sec => sec.classList.remove("active"));
  document.getElementById(`admin-sec-${subTabName}`).classList.add("active");
  
  if (subTabName === "slots") {
    renderAdminSlotsDashboard();
  } else if (subTabName === "unregistered") {
    renderAdminUnregisteredList();
  }
}

function copyReminder(empId) {
  if (!STATE.adminDashboardData) return;
  const { employees } = STATE.adminDashboardData;
  const emp = employees.find(e => e.employeeId === empId);
  if (!emp) return;
  
  const portalUrl = `${window.location.origin}${window.location.pathname}`;
  const text = `แจ้งเตือน: คุณ ${emp.firstName} ${emp.lastName} (รหัสพนักงาน ${emp.employeeId}) แผนก ${emp.department} ยังไม่ได้ลงทะเบียนตรวจสุขภาพประจำปี 2569 รบกวนดำเนินการลงทะเบียนโดยเร็วที่สุดผ่านลิงก์นี้ค่ะ: ${portalUrl}`;
  
  navigator.clipboard.writeText(text)
    .then(() => {
      showToast(`คัดลอกข้อความเตือนคุณ ${emp.firstName} สำเร็จแล้ว`, "success");
    })
    .catch(err => {
      console.error(err);
      showToast("ไม่สามารถคัดลอกข้อความได้", "error");
    });
}

function copyReminderBulk() {
  if (!STATE.adminDashboardData) return;
  const { employees, registrations } = STATE.adminDashboardData;
  const deptFilter = document.getElementById("admin-unreg-dept-select").value;
  
  const unregistered = employees.filter(emp => 
    !registrations.some(r => r.employeeId === emp.employeeId)
  );
  
  const filtered = unregistered.filter(emp => 
    deptFilter === "all" || emp.department === deptFilter
  );
  
  if (filtered.length === 0) {
    showToast("ไม่มีพนักงานค้างลงทะเบียนเพื่อส่งคำเตือน", "warning");
    return;
  }
  
  const portalUrl = `${window.location.origin}${window.location.pathname}`;
  let text = "";
  
  if (deptFilter !== "all") {
    // Message for department representative
    text = `เรียน ตัวแทนแผนก ${deptFilter},\n\nรบกวนช่วยประสานงานติดตามพนักงานที่ยังไม่ได้ลงทะเบียนตรวจสุขภาพประจำปี 2569 จำนวน ${filtered.length} ท่าน ดังรายชื่อด้านล่างนี้:\n`;
    filtered.forEach((emp, idx) => {
      text += `${idx + 1}. รหัส ${emp.employeeId} - คุณ ${emp.firstName} ${emp.lastName}\n`;
    });
    text += `\nรบกวนแจ้งให้พนักงานดำเนินการลงทะเบียนด้วยตนเองผ่านลิงก์นี้: ${portalUrl}\nขอบคุณค่ะ`;
  } else {
    // Message for all departments grouped
    text = `แจ้งเตือนรายชื่อพนักงานที่ยังไม่ได้ลงทะเบียนตรวจสุขภาพประจำปี 2569:\n\n`;
    const depts = [...new Set(filtered.map(e => e.department).filter(Boolean))].sort();
    depts.forEach(dept => {
      const deptEmps = filtered.filter(e => e.department === dept);
      text += `[แผนก ${dept}] (ค้างลงทะเบียน ${deptEmps.length} ท่าน)\n`;
      deptEmps.forEach(emp => {
        text += `- รหัส ${emp.employeeId} : คุณ ${emp.firstName} ${emp.lastName}\n`;
      });
      text += `\n`;
    });
    text += `โปรดแจ้งพนักงานดำเนินการลงทะเบียนตรวจสุขภาพประจำปีผ่านลิงก์นี้: ${portalUrl}\nขอบคุณค่ะ`;
  }
  
  navigator.clipboard.writeText(text)
    .then(() => {
      showToast(`คัดลอกข้อความเตือนของแผนก ${deptFilter === "all" ? "ทั้งหมด" : deptFilter} สำเร็จแล้ว`, "success");
    })
    .catch(err => {
      console.error(err);
      showToast("ไม่สามารถคัดลอกข้อความได้", "error");
    });
}




