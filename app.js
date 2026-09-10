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
  adminDashboardData: null,
  allowCancellation: false,
  isRegistrationClosed: false
};

// --- Client-Side High-Speed Memory & Session Cache (0ms response time) ---
const CLIENT_CACHE = {
  employees: {},
  registrations: {},
  config: null,
  configTimestamp: 0
};

// --- Special Self-Pay Test Catalog Data (34 items) ---
const SPECIAL_TESTS = [
  { id: 1, name: "น้ำตาลในเลือด (FBS)", price: 30, fasting: true, category: "blood", icon: '<i class="fa-solid fa-droplet" style="color: #ef4444;"></i>', notes: "สิทธิประกันสังคมพนักงานอายุ 35 ปีขึ้นไปฟรี (กรณีไม่เคยใช้สิทธิ์ตรวจของประกันสังคมในปีนี้) | *งดอาหารและเครื่องดื่ม 8-12 ชั่วโมง (ดื่มน้ำเปล่าได้)*" },
  { id: 2, name: "ตรวจหาระดับไขมันในเลือด (Cholesterol, Triglyceride, HDL, LDL)", price: 142, fasting: true, category: "blood", icon: '<i class="fa-solid fa-flask" style="color: #f59e0b;"></i>', notes: "งดอาหารและเครื่องดื่ม 8-12 ชั่วโมง (ดื่มน้ำเปล่าได้) | ฟรี สำหรับพนักงานอายุ 35 ปีขึ้นไป (พนักงานอายุไม่ถึง 35 ปี ตรวจเฉพาะ Cholesterol, HDL ฟรี โดยต้องงดอาหารและเครื่องดื่ม)" },
  { id: 3, name: "ตรวจสมรรถภาพการทำงานของตับ (SGO, SGPT, ALK)", price: 150, fasting: false, category: "blood", icon: '<i class="fa-solid fa-heart-pulse" style="color: #10b981;"></i>', notes: "ดูความผิดปกติของตับ ที่อาจทำให้เป็นโรคตับอักเสบ หรือโรคมะเร็งตับ" },
  { id: 4, name: "ตรวจการทำงานของไต (BUN, Creatinine)", price: 60, fasting: false, category: "blood", icon: '<i class="fa-solid fa-stethoscope" style="color: #06b6d4;"></i>', notes: "ตรวจสอบการทำงานของไตว่าทำงานปกติหรือไม่" },
  { id: 5, name: "ตรวจภาวะไทรอยด์ (TFT = FT3, FT4, TSH)", price: 400, fasting: false, category: "blood", icon: '<i class="fa-solid fa-dna" style="color: #8b5cf6;"></i>', notes: "ควบคุมการเผาผลาญ บ่งบอกต่อมไทรอยด์ทำงานผิดปกติหรือไม่" },
  { id: 6, name: "ตรวจหาภาวะโรคเก๊าท์ (Uric Acid)", price: 50, fasting: false, category: "blood", icon: '<i class="fa-solid fa-bone" style="color: #d97706;"></i>', notes: "ดูปริมาณกรดยูริกในเลือดเพื่อวินิจฉัยโรคเก๊าท์" },
  { id: 7, name: "ตรวจหาน้ำตาลสะสม (คัดกรองโรคเบาหวาน) (HbA1C)", price: 300, fasting: false, category: "blood", icon: '<i class="fa-solid fa-chart-line" style="color: #ec4899;"></i>', notes: "วัดระดับน้ำตาลสะสมในเลือดตลอดระยะเวลา 4 เดือนที่ผ่านมา" },
  { id: 8, name: "ตรวจเกลือแร่ในเลือด (Electrolyte)", price: 300, fasting: false, category: "blood", icon: '<i class="fa-solid fa-bolt" style="color: #eab308;"></i>', notes: "ตรวจปริมาณโซเดียม โปแตสเซียม ที่มีผลต่อกล้ามเนื้อหัวใจ" },
  { id: 9, name: "ตรวจหาระดับแคลเซียมในเลือด (Total Calcium)", price: 50, fasting: false, category: "blood", icon: '<i class="fa-solid fa-cubes" style="color: #64748b;"></i>', notes: "ดูปริมาณแคลเซียมในเลือด" },
  { id: 10, name: "ตรวจหาเชื้อไวรัสตับอักเสบชนิดบี (HBs Ag)", price: 100, fasting: false, category: "blood", icon: '<i class="fa-solid fa-shield-virus" style="color: #f43f5e;"></i>', notes: "พนักงานที่เกิดก่อนปี พ.ศ. 2535 ตรวจฟรีโดยใช้สิทธิ์ประกันสังคม (กรณีที่ไม่เคยตรวจ)" },
  { id: 11, name: "ตรวจหาภูมิคุ้มกันเชื้อไวรัสตับอักเสบบี (Anti-HBs)", price: 150, fasting: false, category: "blood", icon: '<i class="fa-solid fa-shield-halved" style="color: #10b981;"></i>', notes: "ตรวจเพื่อดูว่ามีภูมิต้านทานเชื้อไวรัสตับอักเสบ B หรือไม่" },
  { id: 12, name: "ตรวจหาไวรัสตับอักเสบซี (Anti-HCV)", price: 300, fasting: false, category: "blood", icon: '<i class="fa-solid fa-virus" style="color: #8b5cf6;"></i>', notes: "ตรวจหาเชื้อไวรัสตับอักเสบ C ในร่างกาย" },
  { id: 13, name: "ตรวจหากรุ๊ปเลือด (Blood Group)", price: 50, fasting: false, category: "blood", icon: '<i class="fa-solid fa-syringe" style="color: #ef4444;"></i>', notes: "ตรวจระบุหมู่โลหิต (Blood Group)" },
  { id: 14, name: "ตรวจคลื่นไฟฟ้าหัวใจ (EKG)", price: 200, fasting: false, category: "blood", icon: '<i class="fa-solid fa-wave-square" style="color: #dc2626;"></i>', notes: "LPN1: ตรวจที่ห้องพยาบาล Plant 3 (วันที่ 1 และ 7 ต.ค.69) | LPN2: ห้องพยาบาล (วันที่ 2 และ 5 ต.ค.69)" },
  { id: 15, name: "ตรวจคัดกรองมะเร็งต่อมลูกหมาก (PSA)", price: 300, fasting: false, category: "cancer", gender: "M", icon: '<i class="fa-solid fa-person" style="color: #2563eb;"></i>', notes: "สารบ่งชี้มะเร็งต่อมลูกหมากในผู้ชาย" },
  { id: 16, name: "ตรวจคัดกรองมะเร็งตับ (AFP)", price: 300, fasting: false, category: "cancer", icon: '<i class="fa-solid fa-ribbon" style="color: #f59e0b;"></i>', notes: "สารบ่งชี้มะเร็งเพื่อช่วยวินิจฉัยมะเร็งตับ" },
  { id: 17, name: "ตรวจคัดกรองมะเร็งทางเดินอาหาร (มะเร็งลำไส้) (CEA)", price: 300, fasting: false, category: "cancer", icon: '<i class="fa-solid fa-ribbon" style="color: #ea580c;"></i>', notes: "สารบ่งชี้มะเร็งเพื่อช่วยวินิจฉัยมะเร็งทางเดินอาหาร/ลำไส้" },
  { id: 18, name: "ตรวจคัดกรองมะเร็งรังไข่ (CA125)", price: 500, fasting: false, category: "cancer", gender: "F", icon: '<i class="fa-solid fa-ribbon" style="color: #ec4899;"></i>', notes: "สารบ่งชี้มะเร็งรังไข่ในผู้หญิง" },
  { id: 19, name: "ตรวจคัดกรองมะเร็งตับอ่อน (CA 19-9)", price: 500, fasting: false, category: "cancer", icon: '<i class="fa-solid fa-ribbon" style="color: #8b5cf6;"></i>', notes: "สารบ่งชี้มะเร็งตับอ่อนและท่อน้ำดี" },
  { id: 20, name: "ตรวจคัดกรองมะเร็งเต้านม (CA153)", price: 500, fasting: false, category: "cancer", gender: "F", icon: '<i class="fa-solid fa-ribbon" style="color: #f43f5e;"></i>', notes: "สารบ่งชี้มะเร็งเต้านมในผู้หญิง" },
  { id: 21, name: "ตรวจหาร่องรอยเชื้อไวรัสเอดส์ (Anti-HIV)", price: 180, fasting: false, category: "cancer", icon: '<i class="fa-solid fa-shield-heart" style="color: #ef4444;"></i>', notes: "ตรวจหาร่องรอยเชื้อไวรัสเอดส์ (Anti-HIV)" },
  { id: 22, name: "ตรวจคัดกรองโรคธาลัสซีเมีย (HB Typing)", price: 600, fasting: false, category: "cancer", icon: '<i class="fa-solid fa-vial" style="color: #0284c7;"></i>', notes: "ตรวจคัดกรองโรคเลือดจางธาลัสซีเมีย" },
  { id: 23, name: "ตรวจสมรรถภาพการมองเห็น", price: 30, fasting: false, category: "blood", icon: '<i class="fa-solid fa-eye" style="color: #0d9488;"></i>', notes: "ประเมินการมองเห็นชัดเจน" },
  { id: 24, name: "ตรวจสมรรถภาพปอด", price: 30, fasting: false, category: "blood", icon: '<i class="fa-solid fa-lungs" style="color: #0284c7;"></i>', notes: "ตรวจการทำงานและสมรรถภาพของปอด" },
  { id: 25, name: "ตรวจสมรรถภาพการได้ยิน", price: 30, fasting: false, category: "blood", location: "ตรวจที่รถ X-ray", icon: '<i class="fa-solid fa-ear-listen" style="color: #f59e0b;"></i>', notes: "ตรวจที่รถ X-ray" },
  { id: 26, name: "ตรวจหาปริมาณภูมิหลังฉีดวัคซีน Covid-19", price: 800, fasting: false, category: "blood", icon: '<i class="fa-solid fa-shield-virus" style="color: #6366f1;"></i>', notes: "ตรวจหาปริมาณภูมิคุ้มกันโรคโควิด-19" },
  { id: 27, name: "ตรวจภาวะกล้ามเนื้อหัวใจขาดเลือด Troponnin-I (hsTnI)", price: 390, fasting: false, category: "blood", isNew: true, icon: '<i class="fa-solid fa-heart-circle-bolt" style="color: #e11d48;"></i>', notes: "รายการใหม่ New!! ตรวจบ่งชี้ภาวะกล้ามเนื้อหัวใจขาดเลือด" },
  
  // Welfare program items (items 28-34)
  { id: 28, name: "Mammogram", price: 1800, fasting: false, category: "welfare", gender: "F", welfare: { limit: 1500, minAge: 50 }, icon: '<i class="fa-solid fa-person-dress" style="color: #f43f5e;"></i>', notes: "มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท" },
  { id: 29, name: "Thin Prep", price: 1200, fasting: false, category: "welfare", gender: "F", welfare: { limit: 1500, minAge: 50 }, icon: '<i class="fa-solid fa-venus" style="color: #ec4899;"></i>', notes: "มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท" },
  { id: 30, name: "Prostate Screening", price: 300, fasting: false, category: "welfare", gender: "M", welfare: { limit: 1000, minAge: 50 }, icon: '<i class="fa-solid fa-mars" style="color: #2563eb;"></i>', notes: "มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท" },
  { id: 31, name: "CT Calcium Score", price: 5500, fasting: false, category: "welfare", welfare: { limit: 2000, minAge: 50, minLevel: "M4/T5" }, icon: '<i class="fa-solid fa-heart-pulse" style="color: #dc2626;"></i>', notes: "มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท" },
  { id: 32, name: "Colonoscopy", price: 14000, fasting: false, category: "welfare", welfare: { limit: 15000, minAge: 50, minLevel: "M5/T6", isColon: true }, icon: '<i class="fa-solid fa-microscope" style="color: #0284c7;"></i>', notes: "มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท" },
  { id: 33, name: "Colonoscopy + Biopsy", price: 15000, fasting: false, category: "welfare", welfare: { limit: 15000, minAge: 50, minLevel: "M5/T6", isColon: true }, icon: '<i class="fa-solid fa-microscope" style="color: #7c3aed;"></i>', notes: "มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท" },
  { id: 34, name: "MRI Brain (MRI & MRA)", price: 2400, fasting: false, category: "welfare", welfare: { limit: 10000, minAge: 50, minLevel: "M5/T6", isMri: true }, icon: '<i class="fa-solid fa-brain" style="color: #4f46e5;"></i>', notes: "MRI Brain 2,400 ฿ / MRI+MRA 29,000 ฿ (มีสวัสดิการสำหรับพนักงานอายุ 50 ปีขึ้นไป บริษัทช่วยเหลือค่าใช้จ่าย 50% ของราคา ตามเงื่อนไขบริษัท)" }
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
  { employeeId: "003049", firstName: "วิชัย", lastName: "สุขประเสริฐกุล", department: "OPT", defaultLocation: "LPN1", programName: "โปรแกรม MGR", age: 59, gender: "M", programGroup: "โปรแกรม MGR", riskProgram: "", remark: "", ssoApprovedTests: "ตรวจน้ำตาลในเลือด (FBS), ตรวจการทำงานของไต (Cr)" },
  { employeeId: "004148", firstName: "ประภาพร", lastName: "ศรีประดู่", department: "HRDS", defaultLocation: "LPN2", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 57, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "ตรวจการได้ยิน (Audiogram), ตรวจปัสสาวะหาสารเคมี", isPregnant: true, remark: "", ssoApprovedTests: "ตรวจน้ำตาลในเลือด (FBS), ตรวจหาเชื้อไวรัสตับอักเสบชนิดบี (HBs Ag)" },
  { employeeId: "004379", firstName: "ประคอง", lastName: "อ้อยงาม", department: "QM", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 54, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "ตรวจหาสารตะกั่ว", remark: "", ssoApprovedTests: "ตรวจน้ำตาลในเลือด (FBS), ตรวจการทำงานของไต (Cr)" },
  { employeeId: "004766", firstName: "พวงเพชร", lastName: "มณีฉาย", department: "CADT", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 55, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "", remark: "", ssoApprovedTests: "" },
  { employeeId: "005933", firstName: "ระเบียบ", lastName: "ปาละรัตน์", department: "QM", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 60, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "", remark: "ลาป่วยยาว", ssoApprovedTests: "" },
  { employeeId: "006078", firstName: "อดิเรก", lastName: "อ่อนพรม", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 57, gender: "M", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "", remark: "ลาออก", ssoApprovedTests: "" },
  { employeeId: "006125", firstName: "ยุทธนา", lastName: "สุยะนันทน์", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรม MGR", age: 56, gender: "M", programGroup: "โปรแกรม MGR", riskProgram: "ตรวจคลื่นไฟฟ้าหัวใจ (EKG), ตรวจสมรรถภาพปอด", remark: "", ssoApprovedTests: "ตรวจน้ำตาลในเลือด (FBS)" },
  { employeeId: "007860", firstName: "วิไล", lastName: "แสวงศรี", department: "OP4", defaultLocation: "LPN1", programName: "โปรแกรมอายุ 35 ปีขึ้นไป", age: 49, gender: "F", programGroup: "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป", riskProgram: "", remark: "", ssoApprovedTests: "ตรวจไขมันในเลือด (Cholesterol, HDL), ตรวจหาเชื้อไวรัสตับอักเสบชนิดบี (HBs Ag)" },
  { employeeId: "009892", firstName: "สมชาย", lastName: "ม่วงไหม", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรม MGR", age: 56, gender: "M", programGroup: "โปรแกรม MGR", riskProgram: "", remark: "", ssoApprovedTests: "" },
  { employeeId: "010268", firstName: "ยอดธง", lastName: "กรวิรัตน์", department: "OP2S", defaultLocation: "LPN2", programName: "โปรแกรม MGR", age: 55, gender: "M", programGroup: "MGR", riskProgram: "", remark: "อยู่ Hana เกาะกง", ssoApprovedTests: "" },
  { employeeId: "011382", firstName: "เพียงอัมพร", lastName: "องค์วิศิษฐ์", department: "TRF", defaultLocation: "LPN1", programName: "โปรแกรม MGR", age: 57, gender: "F", programGroup: "โปรแกรม MGR", riskProgram: "", remark: "", ssoApprovedTests: "ตรวจน้ำตาลในเลือด (FBS), ตรวจการทำงานของไต (Cr)" },
  { employeeId: "011999", firstName: "ณัฐพงษ์", lastName: "รักเรียน", department: "IT", defaultLocation: "LPN1", programName: "โปรแกรมอายุไม่ถึง 35 ปี", age: 28, gender: "M", programGroup: "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี", riskProgram: "", checkupRight: "ไม่มีสิทธิ์ (อายุงานไม่ถึง 6 เดือน)", remark: "", ssoApprovedTests: "" },
  { employeeId: "012055", firstName: "กมลวรรณ", lastName: "ทองดี", department: "IT", defaultLocation: "LPN1", programName: "โปรแกรมอายุไม่ถึง 35 ปี", age: 33, gender: "F", programGroup: "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี", riskProgram: "", remark: "", ssoApprovedTests: "ตรวจไขมันในเลือด (Cholesterol, HDL)" }
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
  { slotTime: "13:30 - 14:00", limit: 50 },
  { slotTime: "14:00 - 14:30", limit: 50 },
  { slotTime: "14:30 - 15:00", limit: 50 },
  { slotTime: "15:00 - 15:30", limit: 50 },
  { slotTime: "15:30 - 16:00", limit: 50 }
];

// Initialize Mock Registrations locally in LocalStorage
if (!localStorage.getItem("MOCK_REGISTRATIONS")) {
  const initialRegs = [
    { employeeId: "003049", firstName: "วิชัย", lastName: "สุขประเสริฐกุล", department: "OPT", phone: "2201", shift: "ทีม A", location: "LPN1", dateString: "30 กันยายน 2569", timeString: "07:00 - 07:30", cancerTest: "ตรวจคัดกรองมะเร็งต่อมลูกหมาก (PSA)", timestamp: "2026-07-30 08:30:12" },
    { employeeId: "004148", firstName: "ประภาพร", lastName: "ศรีประดู่", department: "HRDS", phone: "1041", shift: "ทีม B", location: "LPN2", dateString: "5 ตุลาคม 2569", timeString: "09:00 - 09:30", cancerTest: "", timestamp: "2026-07-30 09:12:44" }
  ];
  localStorage.setItem("MOCK_REGISTRATIONS", JSON.stringify(initialRegs));
}

// Initialize Allow Cancellation setting in LocalStorage
if (localStorage.getItem("ALLOW_CANCELLATION") === null) {
  localStorage.setItem("ALLOW_CANCELLATION", "false");
}

// --- Initialize App ---
document.addEventListener("DOMContentLoaded", () => {
  setupModeSelector();
  loadConfigAndCounts();
  initRealTimeSync();
  validateFormCompletion();
  renderSpecialCatalogTable();
  
  // Auto-focus status search input on initial page load
  setTimeout(() => {
    const statusInput = document.getElementById("status-emp-id");
    if (statusInput) statusInput.focus();
  }, 200);
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
async function loadConfigAndCounts(forceRefresh = false) {
  if (CONFIG.currentMode === "mock") {
    STATE.configDates = MOCK_CONFIG_DATES;
    STATE.configTimeSlots = MOCK_CONFIG_TIMESLOTS;
    STATE.allowCancellation = localStorage.getItem("ALLOW_CANCELLATION") === "true";
    STATE.isRegistrationClosed = localStorage.getItem("IS_REGISTRATION_CLOSED") === "true";
    
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    STATE.registrationCounts = {};
    regs.forEach(r => {
      const key = `${r.location}|${r.dateString}|${r.timeString}`;
      STATE.registrationCounts[key] = (STATE.registrationCounts[key] || 0) + 1;
    });
    updateRegistrationClosedUI();
  } else {
    // Check in-memory client cache first (valid for 15s) unless forceRefresh requested
    const now = Date.now();
    if (!forceRefresh && CLIENT_CACHE.config && (now - CLIENT_CACHE.configTimestamp < 15000)) {
      STATE.configDates = CLIENT_CACHE.config.dates;
      STATE.configTimeSlots = CLIENT_CACHE.config.timeSlots;
      STATE.registrationCounts = CLIENT_CACHE.config.registrationCounts;
      STATE.allowCancellation = !!CLIENT_CACHE.config.allowCancellation;
      STATE.isRegistrationClosed = !!CLIENT_CACHE.config.isRegistrationClosed;
      updateRegistrationClosedUI();
      return;
    }

    showLoader("กำลังดึงข้อมูลกำหนดการและสิทธิ์การจองล่าสุด...");
    try {
      const response = await callApi("getConfigAndSlots", []);
      if (response && response.success) {
        STATE.configDates = response.data.dates;
        STATE.configTimeSlots = response.data.timeSlots;
        STATE.registrationCounts = response.data.registrationCounts;
        STATE.allowCancellation = !!response.data.allowCancellation;
        STATE.isRegistrationClosed = !!response.data.isRegistrationClosed;
        
        CLIENT_CACHE.config = response.data;
        CLIENT_CACHE.configTimestamp = Date.now();
        
        updateRegistrationClosedUI();
      } else {
        throw new Error(response.error || "ดึงข้อมูลล้มเหลว");
      }
    } catch (err) {
      console.error(err);
      showToast("ไม่สามารถเชื่อมต่อ Google Sheets API ได้ จะใช้ข้อมูลจำลองแทนชั่วคราว", "warning");
      STATE.configDates = MOCK_CONFIG_DATES;
      STATE.configTimeSlots = MOCK_CONFIG_TIMESLOTS;
      STATE.allowCancellation = false;
      STATE.isRegistrationClosed = true; // Default to closed on fallback
      updateRegistrationClosedUI();
    } finally {
      hideLoader();
    }
  }
}

/**
 * Update UI for open/closed registration state
 */
function updateRegistrationClosedUI() {
  const closedBox = document.getElementById("registration-closed-box");
  const regForm = document.getElementById("health-registration-form");
  const isClosed = !!STATE.isRegistrationClosed;
  
  if (closedBox && regForm) {
    if (isClosed) {
      closedBox.style.display = "block";
      regForm.style.display = "none";
    } else {
      closedBox.style.display = "none";
      regForm.style.display = "block";
    }
  }
  
  // If registration is closed and user is currently looking at register tab, auto-switch to status tab
  if (isClosed) {
    const regTabBtn = document.getElementById("tab-btn-register");
    if (regTabBtn && regTabBtn.classList.contains("active")) {
      switchTab("status");
    }
  }
  
  // Update toggle checkbox in Admin Dashboard if rendered
  const adminClosedToggle = document.getElementById("admin-reg-closed-toggle");
  if (adminClosedToggle) {
    adminClosedToggle.checked = isClosed;
  }
  
  // Hide/Show Edit button on Result Card
  const editBtn = document.querySelector("#result-card-container .btn-card-edit");
  if (editBtn) {
    editBtn.style.display = isClosed ? "none" : "inline-flex";
  }
}

// --- Navigation Tabs Control ---
function switchTab(tabName) {
  document.querySelectorAll(".tab-btn").forEach(btn => btn.classList.remove("active"));
  const tabBtn = document.getElementById(`tab-btn-${tabName}`);
  if (tabBtn) {
    tabBtn.classList.add("active");
  }
  
  document.querySelectorAll(".tab-content").forEach(panel => panel.classList.remove("active"));
  const targetPanel = document.getElementById(`tab-${tabName}`);
  if (targetPanel) {
    targetPanel.classList.add("active");
  }
  
  if (tabName === "status") {
    setTimeout(() => {
      const statusInput = document.getElementById("status-emp-id");
      if (statusInput) statusInput.focus();
    }, 50);
  } else if (tabName === "register") {
    setTimeout(() => {
      const regInput = document.getElementById("reg-emp-id");
      if (regInput) regInput.focus();
    }, 50);
  } else if (tabName === "special-catalog") {
    renderSpecialCatalogTable();
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
  
  // 1. Check Fast Client-Side In-Memory & Session Storage Cache (0ms Instant Load)
  if (CONFIG.currentMode === "api") {
    let cachedEmp = CLIENT_CACHE.employees[empId];
    let cachedReg = CLIENT_CACHE.registrations[empId];
    
    if (!cachedEmp) {
      try {
        const stored = sessionStorage.getItem(`emp_${empId}`);
        if (stored) {
          cachedEmp = JSON.parse(stored);
          CLIENT_CACHE.employees[empId] = cachedEmp;
        }
      } catch (e) {}
    }
    if (!cachedReg) {
      try {
        const storedReg = sessionStorage.getItem(`reg_${empId}`);
        if (storedReg) {
          cachedReg = JSON.parse(storedReg);
          CLIENT_CACHE.registrations[empId] = cachedReg;
        }
      } catch (e) {}
    }
    
    // If employee profile is in client cache, render immediately with 0ms delay!
    if (cachedEmp) {
      handleEmployeeLookupResult(cachedEmp, cachedReg || null);
      return;
    }
  }
  
  showLoader("กำลังค้นหาข้อมูลพนักงาน...");
  
  try {
    if (CONFIG.currentMode === "mock") {
      await new Promise(resolve => setTimeout(resolve, 300));
      const emp = MOCK_EMPLOYEES.find(e => e.employeeId === empId);
      if (emp) {
        const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
        const reg = regs.find(r => r.employeeId === empId);
        handleEmployeeLookupResult(emp, reg);
      } else {
        handleEmployeeLookupResult(null, null);
      }
    } else {
      let res = null;
      let isFallbackTriggered = false;
      try {
        res = await callApi("getEmployeeAndRegistration", [empId]);
      } catch (apiErr) {
        console.warn("getEmployeeAndRegistration call failed:", apiErr);
      }
      
      if (res && res.success && res.data) {
        const emp = res.data.employee;
        const reg = res.data.registration;
        
        // Cache to memory and sessionStorage for instant re-access
        if (emp) {
          CLIENT_CACHE.employees[empId] = emp;
          try { sessionStorage.setItem(`emp_${empId}`, JSON.stringify(emp)); } catch (e) {}
        }
        if (reg) {
          CLIENT_CACHE.registrations[empId] = reg;
          try { sessionStorage.setItem(`reg_${empId}`, JSON.stringify(reg)); } catch (e) {}
        }
        
        handleEmployeeLookupResult(emp, reg);
      } else {
        // Fallback: If backend script is not updated yet (throws Action not found)
        const errMsg = (res && res.error) || "";
        if (errMsg.indexOf("Action not found") !== -1 || errMsg.indexOf("not found") !== -1) {
          isFallbackTriggered = true;
          try {
            const empRes = await callApi("getEmployeeData", [empId]);
            if (empRes && empRes.success && empRes.data) {
              let reg = null;
              try {
                const regRes = await callApi("getRegistrationByEmpId", [empId]);
                if (regRes && regRes.success) {
                  reg = regRes.data;
                }
              } catch (e) {
                console.warn("Could not fetch registration in fallback:", e);
              }
              if (empRes.data) {
                CLIENT_CACHE.employees[empId] = empRes.data;
                try { sessionStorage.setItem(`emp_${empId}`, JSON.stringify(empRes.data)); } catch (e) {}
              }
              if (reg) {
                CLIENT_CACHE.registrations[empId] = reg;
                try { sessionStorage.setItem(`reg_${empId}`, JSON.stringify(reg)); } catch (e) {}
              }
              handleEmployeeLookupResult(empRes.data, reg);
              showToast("โปรดอัปเดตสคริปต์ Google Script หลังบ้านเพื่อเปิดใช้งานการค้นหาความเร็วสูง", "warning");
              return;
            }
          } catch (fallbackErr) {
            console.error("Fallback lookup failed:", fallbackErr);
          }
        }
        throw new Error(errMsg || "เกิดข้อผิดพลาดในการดึงข้อมูล");
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
  
  if (employee.checkupRight && employee.checkupRight.indexOf("ไม่มีสิทธิ์") !== -1) {
    showToast("ขออภัย ท่านยังไม่สามารถตรวจสุขภาพประจำปีนี้ได้ เนื่องจากเข้างานยังไม่ครบ 6 เดือน", "error");
    profileBox.style.display = "none";
    STATE.activeEmployee = null;
    return;
  }
  
  if (employee.remark && employee.remark.trim() !== "") {
    showToast(`ขออภัย ไม่พบสิทธิ์การลงทะเบียนตรวจสุขภาพ เนื่องจากมีหมายเหตุในระบบ: "${employee.remark}" (หากมีข้อสงสัยโปรดติดต่อฝ่ายบุคคล)`, "warning");
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

// --- Helper function to determine test fasting requirement ---
function getTestFastingInfo(testName, defaultNpo = false) {
  const lower = (testName || "").toLowerCase();
  
  // Lipid tests (Cholesterol, HDL, LDL, Triglyceride, ไขมัน) or Sugar (FBS, น้ำตาล): Fasting required
  const hasLipid = lower.includes("ไขมัน") || lower.includes("cholesterol") || lower.includes("hdl") || lower.includes("ldl") || lower.includes("triglyceride") || lower.includes("tg");
  const hasFbs = lower.includes("fbs") || lower.includes("น้ำตาล");
  
  if (defaultNpo || hasLipid || hasFbs) {
    return { isNpo: true, note: "งดอาหารและเครื่องดื่ม" };
  }
  
  return { isNpo: false, note: "" };
}

// --- Render checkup list dynamically ---
function renderCheckupList(employee, isPregnant = false) {
  const itemsList = document.getElementById("checkup-items-list");
  if (!itemsList) return;
  itemsList.innerHTML = "";
  
  const programGroup = employee.programGroup;
  const baseTests = PROGRAM_TESTS[programGroup] || [];
  
  // Clone base company tests
  let tests = baseTests.map(t => ({ ...t, isSso: false }));
  
  // Social Security (SSO) tests calculated by age (Registration Tab estimation preview)
  let ssoTests = [];
  const age = Number(employee.age) || 0;
  
  if (age >= 35 || programGroup === "โปรแกรมที่ 1 อายุ 35 ปีขึ้นไป" || programGroup === "โปรแกรม MGR") {
    // For age 35 and above (เกิดก่อน พ.ศ. 2535):
    // If not MGR (MGR already includes FBS & Kidney in base tests)
    if (programGroup !== "โปรแกรม MGR") {
      ssoTests.push({ name: "ตรวจน้ำตาลในเลือด (FBS)", npo: true, isSso: true });
      ssoTests.push({ name: "ตรวจการทำงานของไต (Cr)", npo: false, isSso: true });
    }
    // HBs Ag ได้รับสิทธิ์เฉพาะผู้ที่เกิดก่อน พ.ศ. 2535 (อายุ 35 ปีขึ้นไป ในปี 2569)
    ssoTests.push({ name: "ตรวจหาเชื้อไวรัสตับอักเสบชนิดบี (HBs Ag)", npo: false, isSso: true });
  } else {
    // For age under 35 (โปรแกรมที่ 2 อายุไม่ถึง 35 ปี / เกิดตั้งแต่ พ.ศ. 2535 เป็นต้นไป):
    // สิทธิ์ประกันสังคม: ตรวจไขมันในเลือด (Cholesterol, HDL) สำหรับอายุ 20-34 ปี (ต้องงดอาหารและเครื่องดื่ม)
    if (programGroup === "โปรแกรมที่ 2 อายุไม่ถึง 35 ปี") {
      ssoTests.push({ name: "ตรวจไขมันในเลือด (Cholesterol, HDL)", npo: true, isSso: true });
    }
  }
  
  // Merge estimated SSO tests (avoid duplicates)
  let hasSsoItems = false;
  ssoTests.forEach(ssoItem => {
    let alreadyExists = false;
    tests.forEach(t => {
      if (t.name.toLowerCase().includes(ssoItem.name.toLowerCase()) || ssoItem.name.toLowerCase().includes(t.name.toLowerCase())) {
        alreadyExists = true;
      }
    });
    if (!alreadyExists) {
      tests.push(ssoItem);
      hasSsoItems = true;
    }
  });
  
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
        nameDisplay = `${t.name} *`;
      }
      
      const fastingInfo = getTestFastingInfo(t.name, t.npo);
      
      if (t.isSso) {
        item.className = fastingInfo.isNpo ? "checkup-item npo sso-merged-item" : "checkup-item sso-merged-item";
      } else {
        item.className = fastingInfo.isNpo ? "checkup-item npo" : "checkup-item";
      }
      
      let iconClass = fastingInfo.isNpo ? "fa-solid fa-triangle-exclamation" : "fa-solid fa-circle-check";
      let extraSpan = "";
      if (fastingInfo.isNpo) {
        extraSpan = `<span class="npo-badge">งดอาหาร-เครื่องดื่ม</span>`;
      }
      
      item.innerHTML = `
        <i class="${iconClass}"></i>
        <span>${i + 1}. ${nameDisplay}</span>
        ${extraSpan}
      `;
    }
    itemsList.appendChild(item);
  });
  
  // Show / Hide the blue SSO note below checklist in registration form
  const ssoNote = document.getElementById("checkup-list-note");
  if (ssoNote) {
    ssoNote.innerHTML = `<i class="fa-solid fa-circle-info"></i> * รายการสีน้ำเงินรอเช็คสิทธิ์ประกันสังคม หากเช็กแล้วยังไม่เคยใช้สิทธิ์จะสามารถใช้สิทธิ์ได้ (กรุณานำบัตรประชาชนตัวจริงมาในวันตรวจ)`;
    ssoNote.style.display = hasSsoItems ? "block" : "none";
  }
  
  const nhsoNote = document.getElementById("checkup-nhso-note");
  if (nhsoNote) {
    nhsoNote.style.display = "block";
  }
  
  // Render Custom Risk Factor checkups in a separate card/container
  const riskBox = document.getElementById("risk-tests-box");
  const riskItemsList = document.getElementById("risk-items-list");
  
  if (riskBox && riskItemsList) {
    riskItemsList.innerHTML = ""; // Clear
    const riskItems = employee.riskProgram ? employee.riskProgram.split(',').map(s => s.trim()).filter(Boolean) : [];
    if (riskItems.length > 0) {
      riskBox.style.display = "block";
      riskItems.forEach((riskText, idx) => {
        const item = document.createElement("div");
        item.className = "checkup-item risk-item";
        item.style.background = "#ffffff";
        item.innerHTML = `
          <i class="fa-solid fa-stethoscope" style="color: #ef4444;"></i>
          <span>${idx + 1}. ${riskText}</span>
          <span class="risk-badge" style="background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca;">ปัจจัยเสี่ยง</span>
        `;
        riskItemsList.appendChild(item);
      });
    } else {
      riskBox.style.display = "none";
    }
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
    
    // Exclude slots >= 14:00 on 1st and 5th of October
    const isFirstOrFifth = dateStr.includes("1 ตุลาคม") || dateStr.includes("5 ตุลาคม");
    if (isFirstOrFifth) {
      const match = slotTime.match(/^(\d{2})[.:](\d{2})/);
      if (match) {
        const hour = parseInt(match[1], 10);
        const minute = parseInt(match[2], 10);
        const timeVal = hour * 60 + minute;
        if (timeVal >= 840) { // 14:00 is 14 * 60 = 840 minutes
          return; // Skip rendering this slot
        }
      }
    }
    
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
  
  if (STATE.isRegistrationClosed) {
    showToast("ระบบปิดรับลงทะเบียนและปรับเปลี่ยนรอบเวลาตรวจสุขภาพแล้วค่ะ หากต้องการแก้ไขกรุณาติดต่อฝ่ายบุคคล", "warning");
    return;
  }
  
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
          // Invalidate and update client cache
          CLIENT_CACHE.registrations[payload.employeeId] = payload;
          try { sessionStorage.setItem(`reg_${payload.employeeId}`, JSON.stringify(payload)); } catch (e) {}
          CLIENT_CACHE.configTimestamp = 0; // Force refresh slot counts on next load
          await loadConfigAndCounts(true);
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
  
  // Admin backdoor login bypass
  if (empId === "ad2026") {
    inputEl.value = "";
    switchTab('admin');
    document.getElementById("admin-login-card").style.display = "none";
    document.getElementById("admin-dashboard-content").style.display = "block";
    loadAdminDashboardData();
    showToast("ยินดีต้อนรับผู้ดูแลระบบ เข้าสู่แผงควบคุม", "success");
    return;
  }
  
  if (/^\d+$/.test(empId)) {
    empId = empId.padStart(6, '0');
    inputEl.value = empId;
  }
  
  // 1. Check Fast Client-Side Cache (0ms Instant Status Display)
  if (CONFIG.currentMode === "api") {
    let cachedReg = CLIENT_CACHE.registrations[empId];
    if (!cachedReg) {
      try {
        const stored = sessionStorage.getItem(`reg_${empId}`);
        if (stored) {
          cachedReg = JSON.parse(stored);
          CLIENT_CACHE.registrations[empId] = cachedReg;
        }
      } catch (e) {}
    }
    if (cachedReg) {
      renderStatusCard(cachedReg, empId);
      return; // Instant 0ms load!
    }
  }
  
  showLoader("กำลังค้นหาข้อมูลการลงทะเบียน...");
  
  if (CONFIG.currentMode === "mock") {
    await new Promise(resolve => setTimeout(resolve, 300));
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    const foundReg = regs.find(r => r.employeeId === empId);
    if (foundReg) {
      const emp = MOCK_EMPLOYEES.find(e => e.employeeId === empId);
      if (emp) {
        foundReg.programGroup = emp.programGroup;
        foundReg.age = emp.age;
        foundReg.gender = emp.gender || "";
        foundReg.ssoApprovedTests = emp.ssoApprovedTests || "";
      }
    }
    
    hideLoader();
    renderStatusCard(foundReg, empId);
  } else {
    try {
      let res = null;
      try {
        res = await callApi("getEmployeeAndRegistration", [empId]);
      } catch (e) {
        console.warn("getEmployeeAndRegistration call failed, trying getRegistrationByEmpId fallback:", e);
      }
      
      if (res && res.success && res.data) {
        hideLoader();
        const emp = res.data.employee;
        const reg = res.data.registration;
        
        if (emp) {
          CLIENT_CACHE.employees[empId] = emp;
          try { sessionStorage.setItem(`emp_${empId}`, JSON.stringify(emp)); } catch (e) {}
        }
        if (reg) {
          CLIENT_CACHE.registrations[empId] = reg;
          try { sessionStorage.setItem(`reg_${empId}`, JSON.stringify(reg)); } catch (e) {}
        }
        renderStatusCard(reg, empId);
      } else {
        // Fallback to getRegistrationByEmpId
        const fallbackRes = await callApi("getRegistrationByEmpId", [empId]);
        hideLoader();
        if (fallbackRes && fallbackRes.success) {
          if (fallbackRes.data) {
            CLIENT_CACHE.registrations[empId] = fallbackRes.data;
            try { sessionStorage.setItem(`reg_${empId}`, JSON.stringify(fallbackRes.data)); } catch (e) {}
          }
          renderStatusCard(fallbackRes.data, empId);
        } else {
          throw new Error((fallbackRes && fallbackRes.error) || "ดึงข้อมูลล้มเหลว");
        }
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
  
  // Render test items checklist
  const checklistContainer = document.getElementById("card-tests-list-container");
  checklistContainer.innerHTML = "";
  
  const programGroup = reg.programGroup;
  const baseTests = PROGRAM_TESTS[programGroup] || [];
  
  // Clone base company tests
  let tests = baseTests.map(t => ({ ...t, isSso: false }));
  
  // Parse approved SSO tests from Column M (ssoApprovedTests)
  const ssoRaw = reg.ssoApprovedTests || "";
  let hasSsoItems = false;
  
  if (ssoRaw && ssoRaw.trim() !== "" && ssoRaw.trim() !== "-" && ssoRaw.trim() !== "ไม่มี") {
    const ssoItems = ssoRaw.split(/[,;\n]+/).map(s => s.trim()).filter(Boolean);
    ssoItems.forEach(ssoItemName => {
      let alreadyExists = false;
      tests.forEach(t => {
        if (t.name.toLowerCase().includes(ssoItemName.toLowerCase()) || ssoItemName.toLowerCase().includes(t.name.toLowerCase())) {
          t.isSso = true;
          alreadyExists = true;
          hasSsoItems = true;
        }
      });
      if (!alreadyExists) {
        const fastingInfo = getTestFastingInfo(ssoItemName, false);
        tests.push({
          name: ssoItemName,
          npo: fastingInfo.isNpo,
          isSso: true
        });
        hasSsoItems = true;
      }
    });
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
      
      // Add extra text for fasting items
      let testNameDisplay = `${index + 1}. ${nameDisplay}`;
      const fastingInfo = getTestFastingInfo(t.name, t.npo);
      
      if (fastingInfo.isNpo) {
        if (!testNameDisplay.includes("งดอาหาร") && !testNameDisplay.includes("งดเครื่องดื่ม") && !testNameDisplay.includes("งดน้ำ")) {
          testNameDisplay += " *งดอาหารและเครื่องดื่ม*";
        }
      }
      
      item.innerHTML = `
        <i class="fa-solid fa-check"></i>
        <span>${testNameDisplay}</span>
      `;
    }
    checklistContainer.appendChild(item);
  });
  
  // Render Custom Risk Factor checkups on Ticket in a separate section
  const ticketRiskSection = document.getElementById("card-risk-section");
  const ticketRiskList = document.getElementById("card-risk-list-container");
  
  if (ticketRiskSection && ticketRiskList) {
    ticketRiskList.innerHTML = ""; // Clear
    const riskItems = reg.riskProgram ? reg.riskProgram.split(',').map(s => s.trim()).filter(Boolean) : [];
    if (riskItems.length > 0) {
      ticketRiskSection.style.display = "block";
      riskItems.forEach((riskText, idx) => {
        const item = document.createElement("div");
        item.className = "ticket-test-item ticket-risk-item";
        item.innerHTML = `
          <i class="fa-solid fa-stethoscope" style="color: #ef4444;"></i>
          <span>${idx + 1}. ${riskText}</span>
        `;
        ticketRiskList.appendChild(item);
      });
    } else {
      ticketRiskSection.style.display = "none";
    }
  }
  
  // Append Selected Cancer Checkup if manager or cancerTest has value
  if (reg.cancerTest && reg.cancerTest.trim() !== "" && reg.cancerTest.trim() !== "-" && reg.cancerTest.trim() !== "ไม่มี") {
    const item = document.createElement("div");
    item.className = "ticket-test-item cancer-gold";
    const cancerIdx = tests.length + 1;
    item.innerHTML = `
      <i class="fa-solid fa-crown"></i>
      <span>${cancerIdx}. ${reg.cancerTest}</span>
    `;
    checklistContainer.appendChild(item);
  }
  
  // Show verified / used SSO note on Ticket Card
  const cardSsoVerifiedNote = document.getElementById("card-sso-verified-note");
  if (cardSsoVerifiedNote) {
    if (hasSsoItems) {
      cardSsoVerifiedNote.className = "ticket-sso-verified-note sso-approved";
      cardSsoVerifiedNote.innerHTML = `<i class="fa-solid fa-circle-check" style="color: #0284c7;"></i> <span>* รายการสีน้ำเงินเป็นรายการตรวจหลังเช็คสิทธิ์ประกันสังคมแล้ว (กรุณานำบัตรประชาชนตัวจริงมาในวันตรวจ)</span>`;
      cardSsoVerifiedNote.style.display = "flex";
    } else {
      cardSsoVerifiedNote.className = "ticket-sso-verified-note sso-used";
      cardSsoVerifiedNote.innerHTML = `<i class="fa-solid fa-circle-exclamation" style="color: #ea580c;"></i> <span>* สิทธิ์ตรวจสุขภาพประกันสังคมของท่านถูกใช้ไปแล้ว ดังนั้นท่านจะได้ตรวจโปรแกรมของบริษัทเท่านั้น</span>`;
      cardSsoVerifiedNote.style.display = "flex";
    }
  }
  
  const cardSsoSection = document.getElementById("card-sso-section");
  if (cardSsoSection) {
    cardSsoSection.style.display = "none";
  }
  
  // Hide or show cancel button based on global configuration
  const cancelBtn = cardContainer.querySelector(".btn-card-cancel");
  if (cancelBtn) {
    cancelBtn.style.display = STATE.allowCancellation ? "inline-flex" : "none";
  }
  
  cardContainer.classList.add("visible");
  cardContainer.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// --- Edit Booking Action ---
function editRegistration() {
  if (STATE.isRegistrationClosed) {
    showToast("ระบบปิดรับการแก้ไขรอบเวลาและข้อมูลการลงทะเบียนแล้วค่ะ หากมีความจำเป็นต้องเปลี่ยนแปลงกรุณาติดต่อฝ่ายบุคคล", "warning");
    return;
  }

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
  if (!STATE.allowCancellation) {
    showToast("ระบบยังไม่เปิดให้พนักงานยกเลิกการลงทะเบียนตรวจสุขภาพด้วยตนเองในขณะนี้ค่ะ", "warning");
    return;
  }
  if (!STATE.activeRegistration) return;
  
  const reg = STATE.activeRegistration;
  const confirmCancel = confirm(`คุณต้องการยกเลิกการลงทะเบียนตรวจสุขภาพสำหรับรหัสพนักงาน ${reg.employeeId} หรือไม่?`);
  if (!confirmCancel) return;
  
  // Prompt user for cancellation reason
  const reason = prompt("กรุณาระบุเหตุผลในการยกเลิกการลงทะเบียน (จำเป็น):");
  if (reason === null) return; // User clicked cancel on prompt
  const trimmedReason = reason.trim();
  if (!trimmedReason) {
    showToast("คุณจำเป็นต้องระบุเหตุผลในการยกเลิกด้วยค่ะ", "error");
    return;
  }
  
  showLoader("กำลังยกเลิกการลงทะเบียนของคุณ...");
  
  if (CONFIG.currentMode === "mock") {
    await new Promise(resolve => setTimeout(resolve, 600));
    const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
    
    // Filter out registration
    const updatedRegs = regs.filter(r => r.employeeId !== reg.employeeId);
    localStorage.setItem("MOCK_REGISTRATIONS", JSON.stringify(updatedRegs));
    console.log(`Mock registration cancelled for ${reg.employeeId}. Reason: ${trimmedReason}`);
    
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
        
        const res = await callApi("deleteRegistration", [reg.employeeId, trimmedReason]);
        if (res && res.success) {
          deleted = true;
          delete CLIENT_CACHE.registrations[reg.employeeId];
          try { sessionStorage.removeItem(`reg_${reg.employeeId}`); } catch (e) {}
          CLIENT_CACHE.configTimestamp = 0;
          await loadConfigAndCounts(true);
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

async function callApi(action, args = [], options = {}) {
  if (!CONFIG.apiUrl || CONFIG.apiUrl.includes("your_deployed_url")) {
    showToast("กรุณาติดตั้ง Web App URL ของ Google Sheets ในไฟล์ app.js ก่อน", "error");
    return { success: false, error: "API URL not configured" };
  }
  
  const readActions = [
    "getEmployeeAndRegistration",
    "getConfigAndSlots",
    "getEmployeeData",
    "getRegistrationByEmpId",
    "getAdminDashboardData",
    "prewarmCache"
  ];
  
  const isRead = readActions.includes(action);
  const timeoutMs = options.timeout || 15000;
  const maxRetries = options.retries !== undefined ? options.retries : 2;
  
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    
    try {
      let response;
      if (isRead) {
        // High-Speed HTTP GET for Read Operations (Skips CORS preflight & Apps Script redirection bottlenecks)
        const params = new URLSearchParams();
        params.append("action", action);
        if (args && args.length > 0) {
          params.append("args", JSON.stringify(args));
        }
        
        response = await fetch(`${CONFIG.apiUrl}?${params.toString()}`, {
          method: "GET",
          mode: "cors",
          signal: controller.signal
        });
      } else {
        // HTTP POST for State Modifications
        const payload = {
          action: action,
          args: args
        };
        
        response = await fetch(CONFIG.apiUrl, {
          method: "POST",
          mode: "cors",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify(payload),
          signal: controller.signal
        });
      }
      
      clearTimeout(timeoutId);
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const result = await response.json();
      return result;
    } catch (err) {
      clearTimeout(timeoutId);
      const isLastAttempt = attempt === maxRetries;
      
      if (err.name === 'AbortError') {
        console.warn(`API call timeout (${timeoutMs}ms) for action: ${action}, attempt: ${attempt + 1}`);
        if (isLastAttempt) {
          throw new Error("การเชื่อมต่อไปยัง Google Sheets หมดเวลา (Timeout) กรุณาลองใหม่อีกครั้ง");
        }
      } else {
        console.warn(`API call error on attempt ${attempt + 1} for action ${action}:`, err);
        if (isLastAttempt) {
          throw err;
        }
      }
      
      // Exponential backoff with random jitter before retry: 400ms - 1200ms
      const delay = Math.floor(Math.pow(2, attempt) * 400 + Math.random() * 300);
      await new Promise(res => setTimeout(res, delay));
    }
  }
}

// --- Real-time Synchronisation Manager ---
const SYNC_CONFIG = {
  minIntervalMs: 25000, // 25s - prevents flooding GAS from 1000+ users
  maxIntervalMs: 40000, // 40s with random jitter
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
      STATE.allowCancellation = !!response.data.allowCancellation;
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

// --- Special Self-Pay Catalog Page Table & View Logic ---
let currentCatalogFilter = "all";
let currentCatalogSearch = "";

function renderSpecialCatalogTable() {
  const tbody = document.getElementById("special-catalog-table-body");
  if (!tbody) return;
  
  tbody.innerHTML = "";
  
  const query = currentCatalogSearch.toLowerCase().trim();
  const cat = currentCatalogFilter;
  
  // Show/hide search clear button
  const clearBtn = document.getElementById("catalog-search-clear");
  if (clearBtn) {
    clearBtn.style.display = query ? "block" : "none";
  }
  
  const filtered = SPECIAL_TESTS.filter(item => {
    // 1. Search Query filtering
    if (query) {
      const matchName = (item.name || "").toLowerCase().includes(query);
      const matchNotes = (item.notes || "").toLowerCase().includes(query);
      const matchPrice = String(item.price || "").includes(query);
      if (!matchName && !matchNotes && !matchPrice) return false;
    }
    
    // 2. Category filtering
    if (cat === "all") return true;
    if (cat === "welfare") return !!item.welfare || item.category === "welfare";
    if (cat === "fasting") return !!item.fasting;
    if (cat === "cancer") return item.category === "cancer";
    if (cat === "blood") return item.category === "blood";
    return true;
  });
  
  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="4" style="text-align: center; padding: 40px 20px; color: #94a3b8; font-size: 0.95rem;">
          <i class="fa-solid fa-folder-open" style="font-size: 2.2rem; margin-bottom: 10px; display: block; color: #fecdd3;"></i>
          ไม่พบรายการตรวจสุขภาพที่ตรงกับคำค้นหา
        </td>
      </tr>
    `;
    return;
  }
  
  filtered.forEach(item => {
    const tr = document.createElement("tr");
    
    // Badges
    let badgesHtml = "";
    if (item.isNew) {
      badgesHtml += `<span class="catalog-badge-tag new-item"><i class="fa-solid fa-sparkles"></i> รายการใหม่</span>`;
    }
    if (item.fasting) {
      badgesHtml += `<span class="catalog-badge-tag fasting"><i class="fa-solid fa-utensils"></i> งดอาหาร-เครื่องดื่ม 8-12 ชม.</span>`;
    }
    if (item.welfare) {
      badgesHtml += `<span class="catalog-badge-tag welfare"><i class="fa-solid fa-gift"></i> สวัสดิการ 50%</span>`;
    }
    if (item.id === 1 || item.id === 2 || item.id === 10) {
      badgesHtml += `<span class="catalog-badge-tag sso"><i class="fa-solid fa-shield-halved"></i> มีสิทธิ์ฟรีประกันสังคม</span>`;
    }
    
    // Price display
    let priceText = `${formatPrice(item.price)}`;
    if (item.id === 34) {
      priceText = "2,400 - 29,000";
    }
    
    // Notes & Icons
    let notesText = item.notes || "-";
    let testIcon = item.icon || '<i class="fa-solid fa-notes-medical" style="color: #f43f5e;"></i>';
    
    tr.innerHTML = `
      <td style="text-align: center;">
        <span class="catalog-num-badge">${item.id}</span>
      </td>
      <td>
        <div class="catalog-test-title">
          <span class="catalog-test-icon">${testIcon}</span>
          <span>${item.name}</span>
        </div>
        <div style="margin-top: 5px; margin-left: 35px; display: flex; flex-wrap: wrap; gap: 4px;">
          ${badgesHtml}
        </div>
      </td>
      <td class="catalog-price-val">
        ${priceText} <span style="font-size: 0.76rem; color: #64748b; font-weight: normal;">บาท</span>
      </td>
      <td>
        <span class="catalog-notes-text">${notesText}</span>
      </td>
    `;
    
    tbody.appendChild(tr);
  });
}

function handleCatalogSearch(val) {
  currentCatalogSearch = val;
  renderSpecialCatalogTable();
}

function clearCatalogSearch() {
  const input = document.getElementById("special-catalog-search");
  if (input) {
    input.value = "";
    currentCatalogSearch = "";
    renderSpecialCatalogTable();
    input.focus();
  }
}

function filterCatalogByChip(category, btnElement) {
  currentCatalogFilter = category;
  document.querySelectorAll(".poster-filter-chips .poster-chip").forEach(b => b.classList.remove("active"));
  if (btnElement) btnElement.classList.add("active");
  renderSpecialCatalogTable();
}

function formatPrice(number) {
  return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
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
        const hasNoRight = emp.checkupRight && emp.checkupRight.indexOf("ไม่มีสิทธิ์") !== -1;
        const hasRemark = emp.remark && emp.remark.trim() !== "";
        return !hasNoRight && !hasRemark;
      });
      const registrations = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
      
      STATE.adminDashboardData = {
        employees: eligibleEmployees,
        registrations: registrations
      };
      STATE.allowCancellation = localStorage.getItem("ALLOW_CANCELLATION") === "true";
      STATE.isRegistrationClosed = localStorage.getItem("IS_REGISTRATION_CLOSED") === "true";
    } else {
      const res = await callApi("getAdminDashboardData", []);
      if (res && res.success) {
        STATE.adminDashboardData = res.data;
        STATE.allowCancellation = !!res.data.allowCancellation;
        STATE.isRegistrationClosed = !!res.data.isRegistrationClosed;
      } else {
        throw new Error(res.error || "ไม่สามารถดึงข้อมูลแดชบอร์ดได้");
      }
    }
    renderAdminDashboard();
    updateRegistrationClosedUI();
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
  
  // Set the checkbox states
  const cancelCheckbox = document.getElementById("admin-allow-cancel-toggle");
  if (cancelCheckbox) {
    cancelCheckbox.checked = !!STATE.allowCancellation;
  }

  const closedCheckbox = document.getElementById("admin-reg-closed-toggle");
  if (closedCheckbox) {
    closedCheckbox.checked = !!STATE.isRegistrationClosed;
  }
  
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

async function triggerAutoAllocation() {
  const confirmResult = confirm(
    "หากยืนยัน ระบบจะดำเนินการจัดสรร วันตรวจ และรอบเวลาตรวจ ที่ยังมีที่นั่งว่างอยู่ ให้กับพนักงานทุกคนที่ยังไม่ได้ลงทะเบียนโดยอัตโนมัติ (โดยจะยกเว้นผู้ที่มีหมายเหตุ เช่น ลาออก, ลาป่วยยาว, อยู่ Hana เกาะกง)\n\n" +
    "🔒 และหลังจากจัดสรรเสร็จสิ้น ระบบจะทำการปิดรับการลงทะเบียนและแก้ไขรอบเวลาโดยอัตโนมัติทันที เพื่อป้องกันไม่ให้พนักงานเข้ามาแก้ไขรอบเวลาตรวจ\n\n" +
    "คุณแอดมินยืนยันที่จะดำเนินการหรือไม่?"
  );
  if (!confirmResult) return;
  
  showLoader("กำลังจัดสรรรอบตรวจสุขภาพอัตโนมัติ...");
  
  try {
    let result;
    if (CONFIG.currentMode === "mock") {
      await new Promise(resolve => setTimeout(resolve, 1500));
      // Simulate auto-allocation using mock data
      const regs = JSON.parse(localStorage.getItem("MOCK_REGISTRATIONS") || "[]");
      const unregistered = MOCK_EMPLOYEES.filter(emp => {
        const hasNoRight = emp.checkupRight && emp.checkupRight.indexOf("ไม่มีสิทธิ์") !== -1;
        const hasRemark = emp.remark && emp.remark.trim() !== "";
        const isAlreadyReg = regs.some(r => r.employeeId === emp.employeeId);
        return !hasNoRight && !hasRemark && !isAlreadyReg;
      });
      
      const remarkedEmps = MOCK_EMPLOYEES.filter(emp => {
        const hasRemark = emp.remark && emp.remark.trim() !== "";
        const isAlreadyReg = regs.some(r => r.employeeId === emp.employeeId);
        return hasRemark && !isAlreadyReg;
      });
      
      let successCount = 0;
      let skipCount = remarkedEmps.length;
      let noSlotCount = 0;
      
      unregistered.forEach(emp => {
        successCount++;
        // Add mock registration
        regs.push({
          employeeId: emp.employeeId,
          firstName: emp.firstName,
          lastName: emp.lastName,
          department: emp.department,
          phone: "Auto",
          shift: emp.riskProgram ? "ทีม A" : "คร่อมกะ",
          location: emp.defaultLocation || "LPN1",
          dateString: "30 กันยายน 2569",
          timeString: "07.00 - 07.30 น.",
          cancerTest: "",
          riskProgram: emp.riskProgram || "",
          isPregnant: emp.isPregnant || false,
          ssoConsent: "",
          timestamp: new Date().toISOString()
        });
      });
      
      localStorage.setItem("MOCK_REGISTRATIONS", JSON.stringify(regs));
      localStorage.setItem("IS_REGISTRATION_CLOSED", "true");
      STATE.isRegistrationClosed = true;
      
      result = {
        success: true,
        successCount: successCount,
        skipCount: skipCount,
        noSlotCount: noSlotCount,
        isRegistrationClosed: true
      };
    } else {
      const res = await callApi("autoAllocateRemainingEmployees");
      if (res && res.success) {
        result = res.data;
        STATE.isRegistrationClosed = true;
      } else {
        throw new Error((res && res.error) || "เกิดข้อผิดพลาดในการจัดสรรรอบตรวจ");
      }
    }
    
    hideLoader();
    
    // Update closed state UI immediately
    updateRegistrationClosedUI();
    
    // Show summary alert
    alert(`🎉 ระบบจัดสรรรอบเวลาอัตโนมัติเสร็จสิ้น!\n\n- จัดสรรรอบสำเร็จ: ${result.successCount} ท่าน\n- ข้าม (มีหมายเหตุ ลาออก/ลาป่วยยาว/เกาะกง/คลอด): ${result.skipCount} ท่าน\n- พนักงานที่รอบเต็มไม่มีให้จัดสรร: ${result.noSlotCount} ท่าน\n\n🔒 ระบบได้ทำการปิดรับการลงทะเบียนและแก้ไขรอบเวลาสำหรับพนักงานโดยอัตโนมัติเรียบร้อยแล้วค่ะ`);
    
    // Reload admin statistics
    await loadAdminDashboardData();
    
  } catch (err) {
    console.error(err);
    hideLoader();
    showToast(`เกิดข้อผิดพลาด: ${err.message}`, "error");
  }
}

/**
 * Toggle registration closed setting from Admin Panel
 */
async function toggleRegClosedSetting(checked) {
  if (CONFIG.currentMode === "mock") {
    localStorage.setItem("IS_REGISTRATION_CLOSED", checked ? "true" : "false");
    STATE.isRegistrationClosed = checked;
    showToast(checked ? "🔒 ปิดรับการลงทะเบียนและแก้ไขรอบเวลาแล้ว" : "🔓 เปิดรับการลงทะเบียนเรียบร้อยแล้ว", "success");
    updateRegistrationClosedUI();
  } else {
    showLoader("กำลังอัปเดตการตั้งค่าระบบลงทะเบียน...");
    try {
      const res = await callApi("saveSetting", ["is_registration_closed", checked ? "TRUE" : "FALSE"]);
      if (res && res.success) {
        STATE.isRegistrationClosed = checked;
        showToast(checked ? "🔒 ปิดรับการลงทะเบียนและแก้ไขรอบเวลาแล้ว" : "🔓 เปิดรับการลงทะเบียนเรียบร้อยแล้ว", "success");
        updateRegistrationClosedUI();
      } else {
        throw new Error(res.error || "บันทึกข้อมูลล้มเหลว");
      }
    } catch (err) {
      console.error(err);
      showToast(`มีข้อผิดพลาด: ${err.message}`, "error");
      // Revert checkbox state
      const checkbox = document.getElementById("admin-reg-closed-toggle");
      if (checkbox) {
        checkbox.checked = STATE.isRegistrationClosed;
      }
    } finally {
      hideLoader();
    }
  }
}

/**
 * Toggle allow cancellation setting from Admin Panel
 */
async function toggleCancelButtonSetting(checked) {
  if (CONFIG.currentMode === "mock") {
    localStorage.setItem("ALLOW_CANCELLATION", checked ? "true" : "false");
    STATE.allowCancellation = checked;
    showToast(checked ? "เปิดปุ่มยกเลิกการลงทะเบียนเรียบร้อยแล้วค่ะ" : "ซ่อนปุ่มยกเลิกการลงทะเบียนเรียบร้อยแล้วค่ะ", "success");
    // Update active details card immediately if shown
    const cancelBtn = document.querySelector("#result-card-container .btn-card-cancel");
    if (cancelBtn) {
      cancelBtn.style.display = checked ? "inline-flex" : "none";
    }
  } else {
    showLoader("กำลังอัปเดตการตั้งค่า...");
    try {
      const res = await callApi("saveSetting", ["allow_cancellation", checked ? "TRUE" : "FALSE"]);
      if (res && res.success) {
        STATE.allowCancellation = checked;
        showToast(checked ? "เปิดปุ่มยกเลิกการลงทะเบียนเรียบร้อยแล้วค่ะ" : "ซ่อนปุ่มยกเลิกการลงทะเบียนเรียบร้อยแล้วค่ะ", "success");
        // Update active details card immediately if shown
        const cancelBtn = document.querySelector("#result-card-container .btn-card-cancel");
        if (cancelBtn) {
          cancelBtn.style.display = checked ? "inline-flex" : "none";
        }
      } else {
        throw new Error(res.error || "บันทึกข้อมูลล้มเหลว");
      }
    } catch (err) {
      console.error(err);
      showToast(`มีข้อผิดพลาด: ${err.message}`, "error");
      // Revert checkbox state
      const checkbox = document.getElementById("admin-allow-cancel-toggle");
      if (checkbox) {
        checkbox.checked = STATE.allowCancellation;
      }
    } finally {
      hideLoader();
    }
  }
}

/**
 * Prewarm GAS in-memory cache from Admin Dashboard
 */
async function triggerPrewarmCache() {
  if (CONFIG.currentMode === "mock") {
    showToast("ระบบออฟไลน์ (Mock) ข้อมูลพร้อมใช้งานทันทีอยู่แล้วค่ะ", "info");
    return;
  }
  
  showLoader("กำลังวอร์มแคชความเร็วสูง (Prewarming Cache)...");
  try {
    const res = await callApi("prewarmCache", []);
    hideLoader();
    if (res && res.success) {
      const count = (res.data && res.data.cachedCount) || "ทั้งหมด";
      showToast(`⚡ วอร์มแคชสำเร็จ! โหลดข้อมูลพนักงาน ${count} รายการเข้าสู่ RAM ของระบบแล้ว พนักงานทุกคนจะค้นหาได้ใน 0.2 วินาที`, "success");
    } else {
      throw new Error((res && res.error) || "วอร์มแคชไม่สำเร็จ");
    }
  } catch (err) {
    hideLoader();
    console.error(err);
    showToast(`วอร์มแคชล้มเหลว: ${err.message}`, "error");
  }
}
