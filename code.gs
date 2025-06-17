/**
 * هذا الملف يحتوي على دوال App Script لنظام إدارة أكاديمية غيث.
 * تم بناء الكود ليتوافق مع هيكلة الشيتات الجديدة.
 */

// -----------------------------------------------------------------------------
// تعريفات ثابتة لأسماء الشيتات والأعمدة (تأكد من مطابقتها مع الشيت الفعلي)
// -----------------------------------------------------------------------------
const MASTER_SHEET_NAME = "مواعيد الطلبة"; // اسم الشيت الرئيسي للطلاب (تأكد من هذا الاسم)
const MASTER_STUDENT_ID_COL = 1; // عمود معرف الطالب (A)
const MASTER_STUDENT_NAME_COL = 2; // عمود اسم الطالب (B)
const MASTER_STUDENT_AGE_COL = 3; // عمود السن (C)
const MASTER_STUDENT_NUMBER_COL = 4; // عمود الرقم (D)
const MASTER_SUB_TYPE_COL = 7; // عمود نوع الاشتراك
const MASTER_SYSTEM_COL = 8; // عمود النظام (H)
const MASTER_TOTAL_ATTENDANCE_COL = 9; // عمود عدد الحصص الحاضرة (I) في شيت المشرف
const MASTER_RENEWAL_STATUS_COL = 10; // عمود حالة التجديد (J)
const MASTER_LAST_PAYMENT_DATE_COL = 11; // عمود تاريخ آخر دفع (K) في شيت المشرف
const MASTER_AMOUNT_COL = 12; // عمود المبلغ (L) في شيت المشرف - **هذا هو عمود الدفع**
const MASTER_SUBSCRIPTION_DATE_COL = 13; // عمود تاريخ الاشتراك (M) - **جديد**
const MASTER_FIRST_DATA_ROW = 3; // الصف الذي تبدأ منه بيانات الطلاب في شيت المشرف

const TEACHERS_SHEET_NAME = "المعلمين"; // اسم شيت المعلمين
const TEACHER_ID_COL = 1; // عمود معرف المعلم في شيت المعلمين (A) - **جديد/مُعدل**
const TEACHER_NAME_COL = 2; // عمود اسم المعلم في شيت المعلمين (B) - **جديد/مُعدل**
const TEACHER_SHEET_URL_COL = 5; // عمود رابط شيت المعلم (إذا كان D في شيت المعلمين)
const TEACHER_FIRST_DATA_ROW = 2; // الصف الذي تبدأ منه بيانات المعلمين في شيت المعلمين

// أعمدة شيت المعلم الفردي
const TEACHER_STUDENT_ID_COL_IN_TEACHER_SHEET = 1; // عمود معرف الطالب في شيت المعلم (A)
const TEACHER_ATTENDANCE_COLS_START = 5; // عمود بداية الحضور (F)
const TEACHER_ATTENDANCE_COLS_END = 16; // عمود نهاية الحضور (Q)


// -----------------------------------------------------------------------------
// تعريفات ثابتة لشيت "السجل التاريخي"
// -----------------------------------------------------------------------------
const HISTORY_SHEET_NAME = "السجل التاريخي"; // اسم شيت السجل التاريخي
const HISTORY_STUDENT_ID_COL = 1; // عمود معرف الطالب (A)
const HISTORY_STUDENT_NAME_COL = 2; // عمود اسم الطالب (B)
const HISTORY_STUDENT_AGE_COL = 3; // عمود السن (C)
const HISTORY_STUDENT_NUMBER_COL = 4; // عمود الرقم (D)

const HISTORY_MONTH_TRIPLET_SIZE = 3; // عدد الأعمدة لكل شهر (عدد الحصص، التاريخ، المبلغ)
const HISTORY_MONTH_ATTENDANCE_OFFSET = 0; // الفهرس النسبي لعمود عدد الحصص داخل الثلاثية (0, 1, 2)
const HISTORY_MONTH_DATE_OFFSET = 1;        // الفهرس النسبي لعمود التاريخ داخل الثلاثية
const HISTORY_MONTH_AMOUNT_OFFSET = 2;      // الفهرس النسبي لعمود المبلغ داخل الثلاثية

const HISTORY_FIRST_MONTH_COL_START = 5; // عمود E (مارس التاريخ) - حيث تبدأ أعمدة الشهور
const HISTORY_FIRST_DATA_ROW = 3; // الصف الذي تبدأ منه بيانات الطلاب الفعلية (بعد رؤوس الشهر المدمجة ورؤوس الأعمدة الفرعية)

const MONTH_NAMES_AR = [
  "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
  "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
];

// -----------------------------------------------------------------------------
// تعريفات ثابتة للشيتات الجديدة
// -----------------------------------------------------------------------------

// شيت "مواعيد الطلبة" (سجل المواعيد المحجوزة لكل طالب)
const STUDENT_SCHEDULES_SHEET_NAME = "مواعيد الطلبة";
const STUDENT_SCHEDULE_STUDENT_ID_COL = 1; // A
const STUDENT_SCHEDULE_STUDENT_NAME_COL = 2; // B
const STUDENT_SCHEDULE_AGE_COL = 3; // C 
const STUDENT_SCHEDULE_NUMBER_COL = 4; // D 
const STUDENT_SCHEDULE_TEACHER_NAME_COL = 5; // E 
const STUDENT_SCHEDULE_DAY1_COL = 6; // F 
const STUDENT_SCHEDULE_TIME1_COL = 7; // G 
const STUDENT_SCHEDULE_DAY2_COL = 8; // H 
const STUDENT_SCHEDULE_TIME2_COL = 9; // I 
const STUDENT_SCHEDULE_SUB_TYPE_COL = 10; // J 
const STUDENT_SCHEDULE_SYSTEM_COL = 11; // K 
const STUDENT_SCHEDULE_START_DATE_COL = 12; // L 
const STUDENT_SCHEDULE_FIRST_DATA_ROW = 2;


// شيت "الأرشيف"
const ARCHIVE_SHEET_NAME = "الأرشيف";
// أعمدة الأرشيف ستكون مطابقة لأعمدة مواعيد الطلبة + عمودي الحذف + عمود Master Row
const ARCHIVE_STUDENT_ID_COL = 1; // A
const ARCHIVE_STUDENT_NAME_COL = 2; // B
const ARCHIVE_AGE_COL = 3; // C 
const ARCHIVE_NUMBER_COL = 4; // D 
const ARCHIVE_TEACHER_NAME_COL = 5; // E 
const ARCHIVE_DAY1_COL = 6; // F 
const ARCHIVE_TIME1_COL = 7; // G 
const ARCHIVE_DAY2_COL = 8; // H 
const ARCHIVE_TIME2_COL = 9; // I 
const ARCHIVE_SUB_TYPE_COL = 10; // J 
const ARCHIVE_SYSTEM_COL = 11; // K 
const ARCHIVE_START_DATE_COL = 12; // L 
const ARCHIVE_DATE_ARCHIVED_COL = 13; // M - **العمود الجديد لتاريخ الأرشفة**
const ARCHIVE_REASON_ARCHIVED_COL = 14; // N - **العمود الجديد لسبب الأرشفة**
const ARCHIVE_ORIGINAL_MASTER_ROW_COL = 15; // O - **العمود الجديد لحفظ رقم الصف الأصلي في بيانات الطلبة**
const ARCHIVE_FIRST_DATA_ROW = 2; // الصف الذي تبدأ منه بيانات الأرشيف



// شيت "المواعيد المتاحة للمعلمين" (الشيت اللي في الصورة)
const AVAILABLE_SCHEDULES_SHEET_NAME = "المواعيد المتاحة للمعلمين"; // اسم الشيت (حسب الصورة)
const AVAILABLE_SCHEDULE_TEACHER_ID_COL = 1; // عمود معرف المعلم (A)
const AVAILABLE_SCHEDULE_TEACHER_NAME_COL = 2; // عمود اسم المعلم (B)
const AVAILABLE_SCHEDULE_DAY_COL = 3; // عمود اليوم (C)
const AVAILABLE_SCHEDULE_TIMES_START_COL = 4; // أول عمود يبدأ فيه المواعيد (D)
const AVAILABLE_SCHEDULE_HEADER_ROW = 1; // الصف الذي يحتوي على رؤوس الأيام (السبت، الأحد،...)
const AVAILABLE_SCHEDULE_FIRST_DATA_ROW = 2; // الصف الذي تبدأ فيه بيانات المعلمين الفعلية (T001، هاجر رفعت)
const AVAILABLE_SCHEDULE_ROWS_PER_TEACHER = 7; // عدد الصفوف لكل معلم (لدمج الخلايا: 7 أيام)





// ==============================================================================
// 1. الدوال الأساسية لنقطة الدخول والوظائف المساعدة العامة
// ==============================================================================

/**
 * دالة نقطة الدخول الرئيسية لتطبيق الويب.
 * تقوم بعرض ملف HTML الرئيسي (index.html).
 */
function doGet(e) {
  // عرض ملف index.html الرئيسي فقط
  return HtmlService.createHtmlOutputFromFile('index.html')
      .setTitle('نظام إدارة أكاديمية غيث'); // عنوان يظهر في تبويبة المتصفح
}

/**
 * دالة مساعدة لتمكين الواجهة الأمامية من الحصول على رابط تطبيق الويب الحالي.
 * @returns {string} رابط تطبيق الويب.
 */
function getBaseUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * دالة لتوليد معرف طالب جديد وفريد بناءً على آخر معرف في شيت "الطلاب".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} studentsSheet - شيت "الطلاب".
 * @returns {string} معرف الطالب الجديد (مثال: STD001).
 */
function generateUniqueStudentId(studentsSheet) {
  const lastRow = studentsSheet.getLastRow();
  let lastGeneratedIdNum = 0;
  if (lastRow >= 2) { // نبدأ من الصف الثاني لتخطي العناوين
    const studentIds = studentsSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // العمود A لـ Student ID
    const numericIds = studentIds.flat().map(id => {
      const numPart = String(id).replace('STD', '');
      return parseInt(numPart) || 0;
    }).filter(Number);
    lastGeneratedIdNum = numericIds.length > 0 ? Math.max(...numericIds) : 0;
  }
  return `STD${(lastGeneratedIdNum + 1).toString().padStart(3, '0')}`;
}

/**
 * دالة لتوليد معرف اشتراك جديد وفريد بناءً على آخر معرف في شيت "الاشتراكات الحالية".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} subscriptionsSheet - شيت "الاشتراكات الحالية".
 * @returns {string} معرف الاشتراك الجديد (مثال: SUB001).
 */
function generateUniqueSubscriptionId(subscriptionsSheet) {
  const lastRow = subscriptionsSheet.getLastRow();
  let lastGeneratedIdNum = 0;
  if (lastRow >= 2) { // نبدأ من الصف الثاني لتخطي العناوين
    const subIds = subscriptionsSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // العمود A لـ Subscription ID
    const numericIds = subIds.flat().map(id => {
      const numPart = String(id).replace('SUB', '');
      return parseInt(numPart) || 0;
    }).filter(Number);
    lastGeneratedIdNum = numericIds.length > 0 ? Math.max(...numericIds) : 0;
  }
  return `SUB${(lastGeneratedIdNum + 1).toString().padStart(3, '0')}`;
}

/**
 * دالة لتوليد معرف حضور جديد وفريد بناءً على آخر معرف في شيت "سجل الحضور".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} attendanceSheet - شيت "سجل الحضور".
 * @returns {string} معرف الحضور الجديد (مثال: ATT001).
 */
function generateUniqueAttendanceId(attendanceSheet) {
  const lastRow = attendanceSheet.getLastRow();
  let lastGeneratedIdNum = 0;
  if (lastRow >= 2) { // نبدأ من الصف الثاني لتخطي العناوين
    const attIds = attendanceSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // العمود A لـ Attendance ID
    const numericIds = attIds.flat().map(id => {
      const numPart = String(id).replace('ATT', '');
      return parseInt(numPart) || 0;
    }).filter(Number);
    lastGeneratedIdNum = numericIds.length > 0 ? Math.max(...numericIds) : 0;
  }
  return `ATT${(lastGeneratedIdNum + 1).toString().padStart(3, '0')}`;
}

/**
     * تجلب Teacher ID لمعلم معين من شيت "المعلمين" باستخدام اسمه.
     * @param {string} teacherName - اسم المعلم.
     * @returns {string|null} Teacher ID (من العمود A) أو null إذا لم يتم العثور عليه.
     */
    function getTeacherIdByName(teacherName) {
      const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
      if (!teachersSheet) {
        Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين' في getTeacherIdByName.");
        return null;
      }
      const data = teachersSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1] || '').trim() === String(teacherName).trim()) { // العمود B (مؤشر 1): اسم المعلم
          return String(data[i][0]).trim(); // العمود A (مؤشر 0): Teacher ID
        }
      }
      return null;
    }


/**
     * تجلب اسم المعلم من Teacher ID (العمود A)
     * @param {string} teacherId - Teacher ID.
     * @returns {string|null} اسم المعلم (من العمود B) أو null.
     */
    function getTeacherNameById(teacherId) {
      const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
      if (!teachersSheet) {
        Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين' في getTeacherNameById.");
        return null;
      }
      const data = teachersSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === String(teacherId).trim()) { // العمود A (مؤشر 0): Teacher ID
          return String(data[i][1]).trim(); // العمود B (مؤشر 1): اسم المعلم
        }
      }
      return null;
    }

    /**
     * تجلب قائمة بجميع المعلمين الفريدين من شيت "المعلمين" (للفلترة أو القوائم المنسدلة).
     * @returns {Array<string>} قائمة بأسماء المعلمين الفريدين.
     */
    function getAllTeachersList() {
      const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
      if (!teachersSheet) {
        Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين'.");
        return [];
      }
      const data = teachersSheet.getDataRange().getValues();
      const teachers = new Set(); // لضمان القيم الفريدة

      data.forEach((row, index) => {
        if (index === 0) return; // تخطي صف العناوين
        const teacherName = String(row[1] || '').trim(); // العمود B (مؤشر 1): اسم المعلم
        if (teacherName) {
          teachers.add(teacherName);
        }
      });
      Logger.log("قائمة بجميع المعلمين: " + JSON.stringify(Array.from(teachers)));
      return Array.from(teachers); // تحويل Set إلى Array
    }



/**
 * تجلب تفاصيل باقة معينة من شيت "الباقات".
 * @param {string} packageName - اسم الباقة.
 * @returns {Object|null} كائن يحتوي على تفاصيل الباقة (السعر، عدد الحصص الكلي، نوع الباقة) أو null.
 */
function getPackageDetails(packageName) {
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");
  if (!packagesSheet) {
    Logger.log("خطأ: لم يتم العثور على شيت 'الباقات' في getPackageDetails.");
    return null;
  }
  const data = packagesSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // العمود A: اسم الباقة، العمود B: مدة الحصة، العمود C: عدد الحصص الكلي، العمود D: السعر، العمود E: نوع الباقة
    if (String(data[i][0] || '').trim() === String(packageName).trim()) {
      return {
        'اسم الباقة': String(data[i][0] || '').trim(),
        'مدة الحصة (دقيقة)': data[i][1],
        'عدد الحصص الكلي': data[i][2],
        'السعر': data[i][3],
        'نوع الباقة': String(data[i][4] || '').trim()
      };
    }
  }
  return null;
}

/**
 * دالة مساعدة لتحديد إجمالي عدد الحصص المتاحة بناءً على اسم الباقة.
 * تستخدم عدد الحصص الكلي من شيت الباقات.
 * @param {string} packageName - اسم باقة الاشتراك.
 * @returns {number} إجمالي عدد الحصص للباقة.
 */
function getTotalSessionsForPackage(packageName) {
  const packageDetails = getPackageDetails(packageName);
  if (packageDetails && typeof packageDetails['عدد الحصص الكلي'] === 'number') {
    return packageDetails['عدد الحصص الكلي'];
  }
  Logger.log("تحذير: باقة اشتراك غير معروفة أو لا يوجد بها عدد حصص: " + packageName);
  return 0;
}


/**
 * تقوم بحجز موعد محدد في شيت "المواعيد المتاحة للمعلمين" عن طريق وضع Student ID في الخلية.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - شيت "المواعيد المتاحة للمعلمين".
 * @param {string} teacherId - Teacher ID للمعلم (العمود A).
 * @param {string} day - اليوم (مثلاً: "الأحد").
 * @param {string} timeSlotHeader - رأس عمود الميعاد (مثلاً: "09:00 - 09:30").
 * @param {string} studentId - Student ID للطالب الذي يحجز.
 * @param {string} bookingType - نوع الحجز (مثلاً: "عادي" أو "تجريبي").
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function bookTeacherSlot(sheet, teacherId, day, timeSlotHeader, studentId, bookingType) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // صف العناوين

  let teacherRowIndex = -1;
  let timeSlotColIndex = -1; // مؤشر العمود للميعاد المحدد

  // البحث عن صف المعلم واليوم
  for (let i = 1; i < data.length; i++) {
    // العمود A: Teacher ID، العمود B: اليوم
    if (String(data[i][0] || '').trim() === String(teacherId).trim() && String(data[i][1] || '').trim() === String(day).trim()) {
      teacherRowIndex = i; // الصف في مصفوفة البيانات (0-based)
      break;
    }
  }

  if (teacherRowIndex === -1) {
    return { error: `لم يتم العثور على جدول مواعيد للمعلم ID ${teacherId} في يوم ${day}.` };
  }

  // البحث عن عمود الميعاد المحدد
  for (let i = 2; i < headers.length; i++) { // نبدأ من العمود C (مؤشر 2)
    if (String(headers[i] || '').trim() === String(timeSlotHeader).trim()) {
      timeSlotColIndex = i; // مؤشر العمود في مصفوفة البيانات (0-based)
      break;
    }
  }

  if (timeSlotColIndex === -1) {
    return { error: `لم يتم العثور على عمود الميعاد "${timeSlotHeader}" في شيت 'المواعيد المتاحة للمعلمين'.` };
  }

  // الخلية المراد حجزها
  const targetCell = sheet.getRange(teacherRowIndex + 1, timeSlotColIndex + 1); // +1 لتحويل إلى 1-based

  const currentSlotValue = String(targetCell.getValue() || '').trim();
if (currentSlotValue === timeSlotHeader) { // تأكد أن الموعد متاح (يحتوي على رأس العمود)
    // حجز الميعاد: وضع Student ID في الخلية
    targetCell.setValue(studentId);
} else if (currentSlotValue === studentId) { // لو الطالب بيحاول يحجز نفس الميعاد المحجوز لنفسه (تكرار الاستدعاء)
    return { error: `الميعاد ${day} ${timeSlotHeader} محجوز بالفعل لهذا الطالب.` };
} else if (currentSlotValue !== '') { // لو الخلية فيها قيمة غير رأس العمود وغير Student ID (محجوز لطالب آخر)
    return { error: `الميعاد ${day} ${timeSlotHeader} محجوز بالفعل لطالب آخر.` };
} else { // لو الخلية فارغة (غير متاحة)
    return { error: `الميعاد ${day} ${timeSlotHeader} غير متاح للحجز.` };
}

  // حجز الميعاد: وضع Student ID في الخلية
  targetCell.setValue(studentId);

  // تحديث أعمدة تاريخ الحجز ونوع الحجز
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  // هذه الأعمدة يجب أن تكون موجودة في شيت "المواعيد المتاحة للمعلمين"
  // مثلاً، لو عمود "تاريخ الحجز" هو العمود D (مؤشر 3)
  // ولو عمود "نوع الحجز" هو العمود E (مؤشر 4)
  // sheet.getRange(teacherRowIndex + 1, 4).setValue(today); // تاريخ الحجز (D)
  // sheet.getRange(teacherRowIndex + 1, 5).setValue(bookingType); // نوع الحجز (E)

  return { success: `تم حجز الميعاد بنجاح للطالب ${studentId}.` };
}


/**
 * دالة لتوليد معرف طالب تجريبي جديد وفريد بناءً على آخر معرف في شيت "الطلاب التجريبيون".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} trialStudentsSheet - شيت "الطلاب التجريبيون".
 * @returns {string} معرف الطالب التجريبي الجديد (مثال: TRL001).
 */
function generateUniqueTrialId(trialStudentsSheet) {
  const lastRow = trialStudentsSheet.getLastRow();
  let lastGeneratedIdNum = 0;
  if (lastRow >= 2) { // نبدأ من الصف الثاني لتخطي العناوين
    const trialIds = trialStudentsSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // العمود A لـ Trial ID
    const numericIds = trialIds.flat().map(id => {
      const numPart = String(id).replace('TRL', '');
      return parseInt(numPart) || 0;
    }).filter(Number);
    lastGeneratedIdNum = numericIds.length > 0 ? Math.max(...numericIds) : 0;
  }
  return `TRL${(lastGeneratedIdNum + 1).toString().padStart(3, '0')}`;
}


// ==============================================================================
// 2. الدوال الخاصة بصفحة تسجيل الطلاب (form-page)
// ==============================================================================

/**
     * تجلب بيانات المعلمين ومواعيدهم المتاحة لملء القوائم المنسدلة في نموذج التسجيل.
     * هذه الدالة تجلب المواعيد من شيت "المواعيد المتاحة للمعلمين".
     *
     * @returns {Object} كائن يحتوي على بيانات المعلمين وأيامهم ومواعيدهم المتاحة فقط.
     * { "اسم المعلم": { "اليوم": ["الميعاد1", "الميعاد2"], ... }, ... }
     */
    function getTeacherData() {
      const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
      const teacherAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");

      if (!teachersSheet) {
        Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين'.");
        return { error: "شيت 'المعلمين' غير موجود." };
      }
      if (!teacherAvailableSlotsSheet) {
        Logger.log("خطأ: لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
        return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود." };
      }

      const teacherNamesData = teachersSheet.getDataRange().getValues();
      const availableSlotsData = teacherAvailableSlotsSheet.getDataRange().getValues();
      const headers = availableSlotsData[0]; // رؤوس أعمدة المواعيد في شيت المواعيد

      const teachers = {}; // لـ { "اسم المعلم": { "اليوم": ["الميعاد1", "الميعاد2"], ... } }

      // 1. جمع أسماء المعلمين والـ IDs الخاصة بهم
      const teacherIdToNameMap = new Map();
      for (let i = 1; i < teacherNamesData.length; i++) {
        const teacherId = String(teacherNamesData[i][0] || '').trim(); // العمود A: Teacher ID
        const teacherName = String(teacherNamesData[i][1] || '').trim(); // العمود B: اسم المعلم
        if (teacherId && teacherName) {
          teacherIdToNameMap.set(teacherId, teacherName);
          teachers[teacherName] = {}; // تهيئة كائن المعلم للمواعيد
        }
      }

      // 2. تحديد أعمدة المواعيد الديناميكية من رؤوس الأعمدة
      const timeSlotColumns = [];
      const startColIndexForSlots = 2; // العمود C (مؤشر 2)
      for (let i = startColIndexForSlots; i < headers.length; i++) {
        const header = String(headers[i] || '').trim();
        if (header) {
          timeSlotColumns.push({ index: i, header: header });
        }
      }
      Logger.log("getTeacherData: أعمدة المواعيد المحددة: " + JSON.stringify(timeSlotColumns.map(c => c.header)));


      // 3. المرور على بيانات المواعيد المتاحة للمعلمين لجمع المواعيد الشاغرة (الخلايا الفارغة)
      for (let i = 1; i < availableSlotsData.length; i++) {
        const row = availableSlotsData[i];
        const teacherId = String(row[0] || '').trim(); // العمود A: Teacher ID في شيت المواعيد
        const dayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم في شيت المواعيد

        const teacherName = teacherIdToNameMap.get(teacherId);
        if (teacherName && dayOfWeek) {
          if (!teachers[teacherName][dayOfWeek]) {
            teachers[teacherName][dayOfWeek] = [];
          }

          timeSlotColumns.forEach(colInfo => {
            const slotValue = String(row[colInfo.index] || '').trim(); // قيمة الخلية
            const timeSlotHeader = colInfo.header; // رأس العمود

            // لو الخلية فارغة، الموعد متاح للحجز
            if (slotValue === timeSlotHeader) {
              teachers[teacherName][dayOfWeek].push(timeSlotHeader);
            }
          });
        }
      }
      Logger.log("getTeacherData: البيانات المعالجة للمعلمين: " + JSON.stringify(teachers));
      return teachers;
    }

/**
 * جلب قائمة باقات الاشتراك من شيت "الباقات".
 * @returns {Array<string>} قائمة بأسماء الباقات.
 */
function getSubscriptionPackageList() {
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");
  if (!packagesSheet) {
    Logger.log("خطأ: لم يتم العثور على شيت 'الباقات'.");
    return [];
  }
  const data = packagesSheet.getDataRange().getValues();
  const packageNames = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) {
      packageNames.push(name);
    }
  }
  Logger.log("getSubscriptionPackageList: Returned packages: " + JSON.stringify(packageNames)); // <--- إضافة هذا السطر
  return packageNames;
}

/**
 * جلب قائمة حالات الدفع المتاحة (يمكن أن تكون ثابتة أو من شيت).
 * @returns {Array<string>} قائمة بحالات الدفع.
 */
function getPaymentStatusList() {
  const statuses = [
    "تم الدفع",
    "لم يتم الدفع",
    "تم دفع جزء",
    "حلقة تجريبية",
    "لم يشترك"
  ];
  Logger.log("getPaymentStatusList: Returned statuses: " + JSON.stringify(statuses)); // <--- إضافة هذا السطر
  return statuses;
}


/**
 * حفظ بيانات الطالب الجديد.
 * تقوم بحفظ البيانات في شيت "الطلاب"، وإنشاء اشتراك في شيت "الاشتراكات الحالية"،
 * وتحديث المواعيد المحجوزة في شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {Object} formData - كائن يحتوي على بيانات النموذج المرسلة من الواجهة الأمامية.
 */
function saveData(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName("الطلاب");
  const subscriptionsSheet = ss.getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = ss.getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = ss.getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب'.");
    if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    if (!teachersSheet) throw new Error("لم يتم العثور على شيت 'المعلمين'.");

    const teacherId = getTeacherIdByName(formData.regTeacher);
    if (!teacherId) throw new Error(`لم يتم العثور على Teacher ID للمعلم: ${formData.regTeacher}`);

    const newStudentId = generateUniqueStudentId(studentsSheet);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    let studentBasicStatus = "قيد التسجيل";
    if (formData.regPaymentStatus === "تم الدفع") {
      studentBasicStatus = "مشترك";
    } else if (formData.regPaymentStatus === "حلقة تجريبية") {
      studentBasicStatus = "تجريبي";
    } else if (["لم يشترك", "تم دفع جزء"].includes(formData.regPaymentStatus)) {
      studentBasicStatus = "معلق";
    }

    studentsSheet.appendRow([
      newStudentId,
      formData.regName,
      formData.regAge,
      String(formData.regPhone).trim(),
      "",
      "",
      today,
      studentBasicStatus,
      ""
    ]);
    Logger.log(`✅ تم حفظ الطالب ${formData.regName} (ID: ${newStudentId}) في شيت الطلاب.`);

    const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet);
    const packageName = formData.regSubscriptionPackage;
    const packageDetails = getPackageDetails(packageName);

    let subscriptionRenewalStatus = "تم التجديد";
    let totalClassesAttended = 0;
    let subscriptionType = "شهري";

    if (formData.subscriptionType) {
      subscriptionType = formData.subscriptionType;
    } else if (formData.regPaymentStatus === "حلقة تجريبية") {
      subscriptionRenewalStatus = "تجريبي";
    } else if (packageDetails && packageDetails['نوع الباقة']) {
      subscriptionType = packageDetails['نوع الباقة'];
    }

    let endDate = "";
    const startDate = new Date(today);
    if (subscriptionType === "شهري") {
      startDate.setMonth(startDate.getMonth() + 1);
      endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (subscriptionType === "نصف سنوي") {
      startDate.setMonth(startDate.getMonth() + 6);
      endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    subscriptionsSheet.appendRow([
      newSubscriptionId,                                // A: Subscription ID
      newStudentId,                                     // B: Student ID
      subscriptionType,                                 // ✅ C: نوع الاشتراك
      packageName,                                      // D: اسم الباقة
      teacherId,                                        // E: Teacher ID
      today,                                            // F: تاريخ بداية الاشتراك
      endDate,                                          // G: تاريخ نهاية الاشتراك المتوقع
      totalClassesAttended,                             // H: عدد الحصص الحاضرة
      subscriptionRenewalStatus,                        // I: الحالة التفصيلية للتجديد
      today,                                            // J: تاريخ آخر تجديد
      packageDetails ? packageDetails['السعر'] : 0,     // K: المبلغ الكلي
      formData.regPaymentStatus === "تم الدفع"
        ? (packageDetails ? packageDetails['السعر'] : 0) : 0, // L: المدفوع
      formData.regPaymentStatus === "تم الدفع"
        ? 0 : (packageDetails ? packageDetails['السعر'] : 0), // M: المتبقي
      ""                                                // N: ملاحظات
    ]);
    Logger.log(`✅ تم إنشاء اشتراك (${newSubscriptionId}) للطالب ${newStudentId}.`);

    // 🕒 حجز المواعيد
    const bookedSlotsToUpdate = [];
    const bookingTypeForSlot = formData.regPaymentStatus === "حلقة تجريبية" ? "تجريبي" : "عادي";

    if (formData.regDay1 && formData.regTime1) {
      bookedSlotsToUpdate.push({
        teacherId,
        day: formData.regDay1,
        timeSlotHeader: formData.regTime1,
        studentId: newStudentId,
        bookingType: bookingTypeForSlot
      });
    }
    if (formData.regDay2 && formData.regTime2) {
      bookedSlotsToUpdate.push({
        teacherId,
        day: formData.regDay2,
        timeSlotHeader: formData.regTime2,
        studentId: newStudentId,
        bookingType: bookingTypeForSlot
      });
    }

    bookedSlotsToUpdate.forEach(slot => {
      const result = bookTeacherSlot(
        teachersAvailableSlotsSheet,
        slot.teacherId,
        slot.day,
        slot.timeSlotHeader,
        slot.studentId,
        slot.bookingType
      );
      if (result.error) {
        throw new Error(`تعذر حجز الميعاد ${slot.day} ${slot.timeSlotHeader}: ${result.error}`);
      } else {
        Logger.log(`📌 تم حجز الميعاد ${slot.day} ${slot.timeSlotHeader} للطالب ${newStudentId}`);
      }
    });

    Logger.log("✅ اكتملت عملية حفظ الطالب الجديد بنجاح.");
    return { success: "تم تسجيل الطالب الجديد بنجاح." };

  } catch (e) {
    Logger.log("❌ خطأ في saveData: " + e.message);
    return { error: `فشل تسجيل البيانات: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================================
// 3. الدوال الخاصة بصفحة "كل الطلاب" (All Students Page)
// ==============================================================================

/**
 * تجلب بيانات جميع الطلاب من شيت "الطلاب" و "الاشتراكات الحالية"
 * وتحسب الحصص المتبقية.
 *
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب الموحدة للواجهة الأمامية.
 * أو {Object} كائن خطأ إذا كان شيت "الطلاب" غير موجود.
 */
function getAllStudentsData() {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");


  if (!studentsSheet) {
    Logger.log("خطأ: شيت 'الطلاب' غير موجود لـ getAllStudentsData.");
    return { error: "شيت 'الطلاب' غير موجود." };
  }
  if (!subscriptionsSheet) {
    Logger.log("خطأ: شيت 'الاشتراكات الحالية' غير موجود لـ getAllStudentsData.");
    // يمكن الاستمرار ولكن بدون بيانات الاشتراكات التفصيلية
  }
  if (!teachersSheet) {
    Logger.log("خطأ: شيت 'المعلمين' غير موجود لـ getAllStudentsData.");
  }
  if (!teachersAvailableSlotsSheet) {
    Logger.log("خطأ: شيت 'المواعيد المتاحة للمعلمين' غير موجود لـ getAllStudentsData.");
  }

  const studentData = studentsSheet.getDataRange().getValues();
  const allStudents = [];

  // جلب بيانات الاشتراكات لربطها بـ Student ID
  const subscriptionsMap = new Map(); // key: Student ID, value: { subscriptionDetails }
  if (subscriptionsSheet) {
    const subscriptionsValues = subscriptionsSheet.getDataRange().getValues();
    subscriptionsValues.forEach((row, index) => {
      if (index === 0) return; // تخطي صف العناوين في الاشتراكات
      const studentID = String(row[1] || '').trim(); // العمود B: Student ID في الاشتراكات
      if (studentID) {
        subscriptionsMap.set(studentID, {
          subscriptionId: String(row[0] || '').trim(),      // A
          studentId: String(row[1] || '').trim(),           // B
          subscriptionType: String(row[2] || '').trim(),    // ✅ C: نوع الاشتراك
          packageName: String(row[3] || '').trim(),         // ✅ D: اسم الباقة
          teacherId: String(row[4] || '').trim(),           // ✅ E: Teacher ID
          startDate: row[5],                                // ✅ F
          endDate: row[6],                                  // ✅ G
          attendedSessions: row[7],                  // عدد الحصص الحاضرة (G)
          renewalStatus: String(row[8] || '').trim(), // الحالة التفصيلية للتجديد (H)
          lastRenewalDate: row[9],                   // تاريخ آخر تجديد (I)
          totalSubscriptionAmount: row[10],           // مبلغ الاشتراك الكلي (J)
          paidAmount: row[11],                       // المبلغ المدفوع حتى الآن (K)
          remainingAmount: row[12]                   // المبلغ المتبقي (L)
        });
      }
    });
  }

  // جلب بيانات المعلمين (للحصول على اسم المعلم من Teacher ID)
  const teacherIdToNameMap = new Map();
  if (teachersSheet) {
      const teachersValues = teachersSheet.getDataRange().getValues();
      teachersValues.forEach(row => {
          const teacherId = String(row[0] || '').trim(); // العمود A: Teacher ID
          const teacherName = String(row[1] || '').trim(); // العمود B: اسم المعلم
          if (teacherId) {
              teacherIdToNameMap.set(teacherId, teacherName);
          }
      });
  }

  // جلب بيانات المواعيد المحجوزة لكل طالب (لإظهار اليوم والميعاد الخاص بالطالب)
  const studentBookedSlotsMap = new Map(); // key: Student ID, value: [{day, timeSlotHeader}, {day, timeSlotHeader}]
  if (teachersAvailableSlotsSheet) {
      const availableSlotsValues = teachersAvailableSlotsSheet.getDataRange().getValues();
      const headers = availableSlotsValues[0]; // رؤوس أعمدة المواعيد في شيت المواعيد

      const timeSlotHeaders = [];
      const startColIndexForSlots = 2; // العمود C (مؤشر 2)
      for (let i = startColIndexForSlots; i < headers.length; i++) {
          const header = String(headers[i] || '').trim();
          if (header) {
              timeSlotHeaders.push({ index: i, header: header });
          }
      }

      for (let i = 1; i < availableSlotsValues.length; i++) {
          const row = availableSlotsValues[i];
          const dayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم
          const teacherIdInSlot = String(row[0] || '').trim(); // العمود A: Teacher ID

          timeSlotHeaders.forEach(colInfo => {
              const slotValue = String(row[colInfo.index] || '').trim(); // قيمة الخلية
              const timeSlotHeader = colInfo.header; // رأس العمود

              // لو الخلية فيها Student ID (محجوز)
              if (slotValue.startsWith("STD") || slotValue.startsWith("p ")) {
                  const studentIdInCell = slotValue;
                  if (!studentBookedSlotsMap.has(studentIdInCell)) {
                      studentBookedSlotsMap.set(studentIdInCell, []);
                  }
                  studentBookedSlotsMap.get(studentIdInCell).push({
                      day: dayOfWeek,
                      timeSlotHeader: timeSlotHeader,
                      teacherId: teacherIdInSlot // لتتبع من هو المعلم
                  });
              }
          });
      }
  }


  // المرور على بيانات الطلاب وتجميع كل المعلومات
  studentData.forEach((row, index) => {
    if (index === 0) return; // تخطي صف العناوين في الطلاب

    const studentID = String(row[0] || '').trim(); // العمود A: Student ID
    
    // معلومات الطالب الأساسية
    const studentInfo = {
      rowIndex: index + 1, // رقم الصف في شيت "الطلاب" (1-based)
      studentID: studentID,                        // A
      name: String(row[1] || '').trim(),          // B
      age: row[2],                                // C
      phone: String(row[3] || '').trim(),         // D (رقم الهاتف ولي الأمر)
      studentPhone: String(row[4] || '').trim(),  // E (رقم هاتف الطالب)
      country: String(row[5] || '').trim(),      // F
      registrationDate: row[6] ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd") : '', // G
      basicStatus: String(row[7] || '').trim(),   // H (الحالة الأساسية للطالب)
      notes: String(row[8] || '').trim()          // I
    };

    // دمج بيانات الاشتراك
    const subscriptionDetails = subscriptionsMap.get(studentID);
    if (subscriptionDetails) {
        studentInfo.subscriptionId = subscriptionDetails.subscriptionId;
        studentInfo.packageName = subscriptionDetails.packageName;
        studentInfo.teacherId = subscriptionDetails.teacherId;
        studentInfo.teacherName = teacherIdToNameMap.get(subscriptionDetails.teacherId) || subscriptionDetails.teacherId;
        studentInfo.subscriptionStartDate = subscriptionDetails.startDate ? Utilities.formatDate(subscriptionDetails.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        studentInfo.subscriptionEndDate = subscriptionDetails.endDate ? Utilities.formatDate(subscriptionDetails.endDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        studentInfo.attendedSessions = subscriptionDetails.attendedSessions;
        studentInfo.renewalStatus = subscriptionDetails.renewalStatus;
        studentInfo.lastRenewalDate = subscriptionDetails.lastRenewalDate ? Utilities.formatDate(subscriptionDetails.lastRenewalDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        studentInfo.totalSubscriptionAmount = subscriptionDetails.totalSubscriptionAmount;
        studentInfo.paidAmount = subscriptionDetails.paidAmount;
        studentInfo.remainingAmount = subscriptionDetails.remainingAmount;
    } else {
        // قيم افتراضية إذا لم يكن لديه اشتراك
        studentInfo.subscriptionId = '';
        studentInfo.packageName = '';
        studentInfo.teacherId = '';
        studentInfo.teacherName = '';
        studentInfo.subscriptionStartDate = '';
        studentInfo.subscriptionEndDate = '';
        studentInfo.attendedSessions = 0;
        studentInfo.renewalStatus = 'لم يشترك';
        studentInfo.lastRenewalDate = '';
        studentInfo.totalSubscriptionAmount = 0;
        studentInfo.paidAmount = 0;
        studentInfo.remainingAmount = 0;
    }

    // حساب الحصص المتبقية (إذا كانت الباقة محددة العدد)
    const totalSessions = getTotalSessionsForPackage(studentInfo.packageName); // دالة مساعدة
    let remainingSessions = 'غير متاح';
    if (totalSessions > 0 && typeof studentInfo.attendedSessions === 'number') {
        remainingSessions = totalSessions - studentInfo.attendedSessions;
    } else if (studentInfo.renewalStatus === "تجريبي") {
        remainingSessions = "تجريبي";
    } else if (studentInfo.packageName.includes("نصف سنوي")) { // لو باقة نصف سنوية
        remainingSessions = "نصف سنوي";
    }
    studentInfo.remainingSessions = remainingSessions;


    // دمج المواعيد المحجوزة (اليوم الأول والميعاد الأول، إلخ)
    const bookedSlots = studentBookedSlotsMap.get(studentID) || [];
    studentInfo.day1 = bookedSlots[0] ? bookedSlots[0].day : '';
    studentInfo.time1 = bookedSlots[0] ? bookedSlots[0].timeSlotHeader : '';
    studentInfo.day2 = bookedSlots[1] ? bookedSlots[1].day : '';
    studentInfo.time2 = bookedSlots[1] ? bookedSlots[1].timeSlotHeader : '';
    
    // يمكن هنا إضافة Teacher ID للطالب في المواعيد عشان لو فيه أكثر من معلم لنفس الطالب
    // studentInfo.teacherIdForSlot1 = bookedSlots[0] ? bookedSlots[0].teacherId : '';
    // studentInfo.teacherIdForSlot2 = bookedSlots[1] ? bookedSlots[1].teacherId : '';

    allStudents.push(studentInfo);
  });

  Logger.log("تم جلب " + allStudents.length + " طالب لصفحة 'كل الطلاب'.");
  return allStudents;
}

// ==============================================================================
// 5. الدوال الخاصة بلوحة التحكم (Dashboard Page)
// ==============================================================================

/**
 * تجلب كافة الإحصائيات المطلوبة لعرضها في لوحة التحكم (الداشبورد).
 * تجمع البيانات من شيتات "الطلاب" و "الاشتراكات الحالية" و "المعلمين".
 *
 * @returns {Object} كائن يحتوي على جميع الإحصائيات.
 * أو {Object} كائن خطأ إذا كانت الشيتات غير موجودة.
 */
function getDashboardStats() {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");

  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود." };
  if (!subscriptionsSheet) return { error: "شيت 'الاشتراكات الحالية' غير موجود." };
  if (!teachersSheet) return { error: "شيت 'المعلمين' غير موجود." };

  const studentData = studentsSheet.getDataRange().getValues();
  const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
  const teachersData = teachersSheet.getDataRange().getValues();

  const students = studentData.slice(1);
  const subscriptions = subscriptionsData.slice(1);
  const teachers = teachersData.slice(1);

  const stats = {
    totalStudents: students.length,
    registeredStudents: 0,
    trialStudents: 0,
    pendingStudents: 0,
    studentsByPackage: {},
    renewalStatus: {
      needsRenewal: 0,
      expired: 0,
      overLimit: 0,
      renewed: 0,
      trial: 0,
      notSubscribed: 0
    },
    studentsByTeacher: {},
    recentlyRegistered: {
      last7Days: 0,
      last30Days: 0
    },
    renewalNextWeek: 0,
    renewalNextWeekStudents: []
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const nextWeek = new Date(today);
  nextWeek.setDate(today.getDate() + 7);

  const teacherIdToNameMap = new Map();
  teachers.forEach(row => {
    const teacherId = String(row[0] || '').trim();
    const teacherName = String(row[1] || '').trim();
    if (teacherId) {
      teacherIdToNameMap.set(teacherId, teacherName);
    }
  });

  students.forEach(row => {
    const registrationDate = row[6];
    const basicStatus = String(row[7] || '').trim();

    if (basicStatus === "مشترك") stats.registeredStudents++;
    else if (basicStatus === "تجريبي") stats.trialStudents++;
    else if (basicStatus === "قيد التسجيل" || basicStatus === "معلق") stats.pendingStudents++;

    if (registrationDate instanceof Date) {
      const regDate = new Date(registrationDate);
      regDate.setHours(0, 0, 0, 0);
      const diffDays = Math.ceil(Math.abs(today - regDate) / (1000 * 60 * 60 * 24));
      if (diffDays <= 7) stats.recentlyRegistered.last7Days++;
      if (diffDays <= 30) stats.recentlyRegistered.last30Days++;
    }
  });

  subscriptions.forEach(row => {
    const studentId = String(row[1] || '').trim();            // B
    const packageName = String(row[3] || '').trim();           // ✅ C→D
    const teacherId = String(row[4] || '').trim();             // ✅ D→E
    const renewalStatus = String(row[8] || '').trim();         // ✅ H→I
    const lastRenewalDate = row[9];                            // ✅ I→J

    if (renewalStatus === "يحتاج للتجديد") stats.renewalStatus.needsRenewal++;
    else if (renewalStatus === "انتهت مدة الاشتراك") stats.renewalStatus.expired++;
    else if (renewalStatus === "تجاوز الحد") stats.renewalStatus.overLimit++;
    else if (renewalStatus === "تم التجديد") stats.renewalStatus.renewed++;
    else if (renewalStatus === "تجريبي") stats.renewalStatus.trial++;
    else if (renewalStatus === "لم يشترك") stats.renewalStatus.notSubscribed++;

    if (packageName) {
      stats.studentsByPackage[packageName] = (stats.studentsByPackage[packageName] || 0) + 1;
    }

    if (teacherId) {
      const teacherName = teacherIdToNameMap.get(teacherId) || teacherId;
      if (!stats.studentsByTeacher[teacherName]) {
        stats.studentsByTeacher[teacherName] = new Set();
      }
      stats.studentsByTeacher[teacherName].add(studentId);
    }

    if (lastRenewalDate instanceof Date && renewalStatus === "تم التجديد") {
      const renewalCheckDate = new Date(lastRenewalDate);
      renewalCheckDate.setMonth(renewalCheckDate.getMonth() + 1);
      renewalCheckDate.setHours(0, 0, 0, 0);

      if (renewalCheckDate >= today && renewalCheckDate <= nextWeek) {
        stats.renewalNextWeek++;
        const studentRow = students.find(s => String(s[0]).trim() === studentId);
        if (studentRow) {
          stats.renewalNextWeekStudents.push({
            name: String(studentRow[1] || '').trim(),
            phone: String(studentRow[3] || '').trim()
          });
        }
      }
    }
  });

  for (const teacher in stats.studentsByTeacher) {
    stats.studentsByTeacher[teacher] = stats.studentsByTeacher[teacher].size;
  }

  Logger.log("Dashboard Stats: " + JSON.stringify(stats));
  return stats;
}


// ==============================================================================
// 6. الدوال الخاصة بصفحة سجل الحضور للمعلمين (Teacher Schedule Page)
// ==============================================================================

/**
 * تجلب اسم Teacher ID واسم المعلم ورقم هاتفه من شيت "المعلمين" باستخدام رقم الهاتف.
 * @param {string} phone - رقم الهاتف للبحث عنه.
 * @returns {Object|null} كائن يحتوي على { teacherId, teacherName, phone } أو null.
 */
function getTeacherInfoByPhone(phone) { // تم تغيير الاسم من getTeacherNameByPhone ليكون أكثر شمولاً
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  if (!teachersSheet) {
    Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين' في getTeacherInfoByPhone.");
    return { error: "شيت 'المعلمين' غير موجود." };
  }
  const data = teachersSheet.getDataRange().getValues();

  // إزالة أول رقم إذا كان 0 من الهاتف المدخل
  const formattedPhone = String(phone).trim();
  const cleanedSearchPhone = formattedPhone.startsWith("0") ? formattedPhone.substring(1) : formattedPhone;

  Logger.log("البحث عن رقم الهاتف (بعد التنسيق): " + cleanedSearchPhone);

  for (let i = 1; i < data.length; i++) {
    const teacherId = String(data[i][0] || '').trim(); // العمود A: Teacher ID
    const teacherName = String(data[i][1] || '').trim(); // العمود B: اسم المعلم
    const storedPhone = String(data[i][2] || '').trim(); // العمود C: رقم الهاتف

    // تنظيف رقم الهاتف المخزن بنفس الطريقة
    const cleanedStoredPhone = storedPhone.startsWith("0") ? storedPhone.substring(1) : storedPhone;

    if (cleanedStoredPhone === cleanedSearchPhone) {
      Logger.log(`تم العثور على المعلم: ${teacherName} (ID: ${teacherId})`);
      return { teacherId: teacherId, teacherName: teacherName, phone: storedPhone };
    }
  }

  Logger.log("رقم الهاتف غير موجود للمعلم.");
  return null;
}

/**
 * تجلب قائمة الطلاب الذين يجب أن يحضروا لمعلم معين في يوم معين.
 * (تعتمد على المواعيد المحجوزة في شيت "المواعيد المتاحة للمعلمين").
 *
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} day - اليوم (مثلاً: "الأحد").
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب (Student ID, name, timeSlotHeader, bookingType).
 */
function getAttendanceStudentsForTeacherAndDay(teacherId, day) {
    const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
    const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون"); // شيت الطلاب التجريبيين
    const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
    const attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور");

    if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود." };
    if (!trialStudentsSheet) return { error: "شيت 'الطلاب التجريبيون' غير موجود." }; // التحقق
    if (!teachersAvailableSlotsSheet) return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود." };

    // جلب بيانات الطلاب (ID -> Name) من كلا الشيتين
    const studentIdToNameMap = new Map();
  const studentData = studentsSheet.getDataRange().getValues();
  studentData.forEach(row => {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (id) studentIdToNameMap.set(id, name);
  });
  // جديد: جلب أسماء الطلاب التجريبيين أيضاً
  const trialStudentsData = studentsSheet.getDataRange().getValues(); // تأكد من استخدام الشيت الصحيح هنا!
  // خطأ محتمل: يجب أن يكون trialStudentsSheet وليس studentsSheet
  // يجب أن يكون: const trialStudentsData = trialStudentsSheet.getDataRange().getValues();
  if (trialStudentsSheet) { // تأكد من وجود الشيت قبل محاولة قراءته
      const actualTrialStudentsData = trialStudentsSheet.getDataRange().getValues();
      actualTrialStudentsData.forEach(row => {
          const id = String(row[0] || '').trim(); // العمود A: Trial ID
          const name = String(row[1] || '').trim(); // العمود B: Student Name
          if (id) studentIdToNameMap.set(id, name);
      });
  }


    // جلب بيانات الحضور المسجلة لليوم الحالي لمنع تكرار التسجيل في الواجهة
    const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const markedAttendanceToday = new Set(); // Key: studentID_timeSlotHeader
    if (attendanceLogSheet) {
        const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
        attendanceLogData.forEach(row => {
            const logDateValue = row[4]; // العمود E: تاريخ الحصة في سجل الحضور (index 4)
            const logDate = (logDateValue instanceof Date) ? Utilities.formatDate(logDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
            const logStudentID = String(row[1] || '').trim();
            const logTimeSlot = String(row[5] || '').trim();

            if (logDate === todayFormatted) {
                markedAttendanceToday.add(`<span class="math-inline">\{logStudentID\}\_</span>{logTimeSlot}`);
            }
        });
    }

    const studentsForAttendance = [];
    const teacherSlotsData = teachersAvailableSlotsSheet.getDataRange().getValues();
    const headers = teacherSlotsData[0];

    const timeSlotCols = [];
    const startColIndexForSlots = 2;
    for (let i = startColIndexForSlots; i < headers.length; i++) {
        const header = String(headers[i] || '').trim();
        if (header) timeSlotCols.push({ index: i, header: header });
    }

    for (let i = 1; i < teacherSlotsData.length; i++) {
        const row = teacherSlotsData[i];
        const currentTeacherId = String(row[0] || '').trim();
        const currentDayOfWeek = String(row[1] || '').trim();

        if (currentTeacherId === teacherId && currentDayOfWeek === day) {
            timeSlotCols.forEach(colInfo => {
                const slotValue = String(row[colInfo.index] || '').trim();
                const timeSlotHeader = colInfo.header;

                // إذا كانت الخلية تحتوي على Student ID أو Trial ID (محجوزة لطالب)
                if (slotValue.startsWith("STD") || slotValue.startsWith("TRL") || slotValue.startsWith("p ")) { // p للحجوزات القديمة غير معرفة
                    const studentIdInCell = slotValue;
                    const studentName = studentIdToNameMap.get(studentIdInCell) || studentIdInCell;

                    const isMarked = markedAttendanceToday.has(`<span class="math-inline">\{studentIdInCell\}\_</span>{timeSlotHeader}`);

                    studentsForAttendance.push({
                        studentID: studentIdInCell,
                        name: studentName,
                        timeSlot: timeSlotHeader,
                        isMarked: isMarked
                    });
                }
            });
        }
    }

    Logger.log(`تم جلب ${studentsForAttendance.length} طالب حضور للمعلم ${teacherId} في يوم ${day}.`);
    return studentsForAttendance;
}

/**
 * تسجل حضور طالب معين في "سجل الحضور" وتحدث "الاشتراكات الحالية".
 *
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} studentId - Student ID للطالب.
 * @param {string} day - يوم الحصة (مثلاً: "الأحد").
 * @param {string} timeSlot - وقت الحصة (رأس الميعاد: "09:00 - 09:30").
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function markAttendance(teacherId, studentId, day, timeSlot) {
    const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
    const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");
    const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
    let attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور");

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);

        if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
        if (!trialStudentsSheet) throw new Error("شيت 'الطلاب التجريبيون' غير موجود.");
        if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");

        if (!attendanceLogSheet) {
            attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("سجل الحضور");
            attendanceLogSheet.appendRow([
                "Attendance ID", "Student ID", "Teacher ID", "Subscription ID",
                "تاريخ الحصة", "وقت الحصة", "اليوم", "حالة الحضور", "نوع الحصة", "ملاحظات المعلم"
            ]);
        }

        const today = new Date();
        const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

        const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
        for (let i = 1; i < attendanceLogData.length; i++) {
            const logRow = attendanceLogData[i];
            const logStudentID = String(logRow[1] || '').trim();
            const logTeacherID = String(logRow[2] || '').trim();
            const logDateValue = logRow[4];
            const logDate = (logDateValue instanceof Date) ? Utilities.formatDate(logDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
            const logTimeSlot = String(logRow[5] || '').trim();

            if (logStudentID === studentId && logTeacherID === teacherId && logDate === todayFormatted && logTimeSlot === timeSlot) {
                return { error: "تم تسجيل الحضور لهذا الطالب في هذا الميعاد وهذا اليوم مسبقًا." };
            }
        }

        let studentName = '';
        let isTrialStudent = studentId.startsWith("TRL");
        let subscriptionId = '';
        let packageName = '';
        let renewalStatus = '';
        let totalPackageSessions = 0;
        let subscriptionRowIndex = -1;
        let classType = "عادية";

        if (isTrialStudent) {
            const trialStudentsData = trialStudentsSheet.getDataRange().getValues();
            const trialStudentRow = trialStudentsData.find(row => String(row[0] || '').trim() === studentId);
            if (trialStudentRow) {
                studentName = String(trialStudentRow[1] || '').trim();
                classType = "تجريبية";
                Logger.log(`تسجيل حضور لطالب تجريبي: ${studentName}`);
            } else {
                throw new Error(`لم يتم العثور على اسم الطالب التجريبي بمعرف ${studentId}.`);
            }
        } else {
            const studentsData = studentsSheet.getDataRange().getValues();
            const regularStudentRow = studentsData.find(row => String(row[0] || '').trim() === studentId);
            if (regularStudentRow) {
                studentName = String(regularStudentRow[1] || '').trim();
                Logger.log(`تسجيل حضور لطالب مشترك: ${studentName}`);
            } else {
                throw new Error(`لم يتم العثور على اسم الطالب المشترك بمعرف ${studentId}.`);
            }

            const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
            for (let i = 1; i < subscriptionsData.length; i++) {
                if (String(subscriptionsData[i][1] || '').trim() === studentId) {
                    subscriptionRowIndex = i;
                    subscriptionId = String(subscriptionsData[i][0] || '').trim();      // A
                    packageName = String(subscriptionsData[i][3] || '').trim();        // ✅ D بعد الإزاحة
                    renewalStatus = String(subscriptionsData[i][8] || '').trim();      // ✅ I بعد الإزاحة
                    totalPackageSessions = getTotalSessionsForPackage(packageName);
                    break;
                }
            }
            if (subscriptionRowIndex === -1) throw new Error(`لم يتم العثور على اشتراك للطالب ID ${studentId}.`);
        }

        const attendanceId = generateUniqueAttendanceId(attendanceLogSheet);
        attendanceLogSheet.appendRow([
            attendanceId,
            studentId,
            teacherId,
            subscriptionId,
            today,
            timeSlot,
            day,
            "حضر",
            classType,
            ""
        ]);
        Logger.log(`تم تسجيل الحضور لـ ${studentName} (ID: ${studentId}) في ${day} ${timeSlot} كحصة ${classType}.`);

        if (classType === "عادية") {
            const currentAttendedSessionsCell = subscriptionsSheet.getRange(subscriptionRowIndex + 1, 8); // ✅ H = col 8
            let currentSessions = currentAttendedSessionsCell.getValue();
            currentSessions = (typeof currentSessions === 'number') ? currentSessions : 0;
            subscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue(currentSessions + 1);

            Logger.log(`تم تحديث عدد الحصص الحاضرة للطالب ${studentId} إلى ${currentSessions + 1}.`);

            if (totalPackageSessions > 0 && (currentSessions + 1) >= totalPackageSessions) {
                subscriptionsSheet.getRange(subscriptionRowIndex + 1, 9).setValue("يحتاج للتجديد"); // ✅ I = col 9
                Logger.log(`حالة تجديد الطالب ${studentId} تم تحديثها إلى "يحتاج للتجديد".`);
            }
        } else if (classType === "تجريبية") {
            const trialStudentRowIndex = trialStudentsSheet.getDataRange().getValues().findIndex(row => String(row[0] || '').trim() === studentId) + 1;
            if (trialStudentRowIndex > 0) {
                Logger.log(`تم تسجيل حصة تجريبية للطالب ${studentId}.`);
            }
        }

        return { success: `تم تسجيل الحضور بنجاح للطالب ${studentName}.` };

    } catch (e) {
        Logger.log("خطأ في markAttendance: " + e.message);
        return { error: `فشل تسجيل الحضور: ${e.message}` };
    } finally {
        lock.releaseLock();
    }
}



// ==============================================================================
// 7. الدوال الخاصة بصفحة إدارة مواعيد المعلمين (Manage Slots Page) - ستضاف لاحقاً
// ==============================================================================


/**
 * تجلب كافة المواعيد (المتاحة والمحجوزة) لمعلم معين من شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {string} teacherName - اسم المعلم المراد جلب مواعيده.
 * @returns {Object} كائن يحتوي على اسم المعلم، Teacher ID، ومصفوفة من كائنات المواعيد، أو كائن خطأ.
 */
function getTeacherAvailableSlots(teacherName) {
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");

  if (!teachersSheet) return { error: "شيت 'المعلمين' غير موجود." };
  if (!teachersAvailableSlotsSheet) return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود." };
  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود." };

  const teacherId = getTeacherIdByName(teacherName);
  if (!teacherId) {
    return { error: `لم يتم العثور على Teacher ID للمعلم: ${teacherName}` };
  }

  const availableSlotsData = teachersAvailableSlotsSheet.getDataRange().getValues();
  const headers = availableSlotsData[0]; // رؤوس أعمدة المواعيد في شيت المواعيد

  const studentIdToNameMap = new Map();
  const studentData = studentsSheet.getDataRange().getValues();
  studentData.forEach(row => {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (id) studentIdToNameMap.set(id, name);
  });

  const slots = [];
  const timeSlotColumns = [];
  const startColIndexForSlots = 2; // العمود C (مؤشر 2)
  for (let i = startColIndexForSlots; i < headers.length; i++) {
    const header = String(headers[i] || '').trim();
    if (header) {
      timeSlotColumns.push({ index: i, header: header });
    }
  }

  // البحث عن صفوف المعلم المحددة في شيت المواعيد
  for (let i = 1; i < availableSlotsData.length; i++) {
    const row = availableSlotsData[i];
    const currentTeacherId = String(row[0] || '').trim();
    const dayOfWeek = String(row[1] || '').trim();

    if (currentTeacherId === teacherId) { // وجدنا صفوف هذا المعلم
      timeSlotColumns.forEach(colInfo => {
        const slotValue = String(row[colInfo.index] || '').trim();
        const timeSlotHeader = colInfo.header;
        let isBooked = false;
        let bookedBy = null;

        if (slotValue.startsWith("STD") || slotValue.startsWith("TRL")) { // محجوزة بواسطة طالب
          isBooked = true;
          bookedBy = {
            _id: slotValue,
            name: studentIdToNameMap.get(slotValue) || 'طالب غير معروف'
          };
        } else if (slotValue !== '' && slotValue !== timeSlotHeader) { // محجوزة بنص مخصص (مثل "محجوز")
          isBooked = true;
          bookedBy = {
            name: slotValue // النص المخصص هو اسم الحجز
          };
        }

        slots.push({
          dayOfWeek: dayOfWeek,
          timeSlotHeader: timeSlotHeader, // رأس العمود (09:00 - 09:30)
          actualSlotValue: slotValue, // القيمة الفعلية في الخلية
          isBooked: isBooked,
          bookedBy: bookedBy
        });
      });
    }
  }

  Logger.log(`تم جلب ${slots.length} موعد للمعلم ${teacherName}.`);
  return { teacherName: teacherName, teacherId: teacherId, slots: slots };
}


/**
 * تقوم بتحديث المواعيد المتاحة لمعلم في يوم معين في شيت "المواعيد المتاحة للمعلمين".
 * يتم وضع رأس العمود في الخلايا المختارة (لجعلها متاحة) وإفراغ الخلايا غير المختارة (لجعلها غير متاحة).
 * لا يمكن تعديل المواعيد المحجوزة بالفعل.
 *
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} day - اليوم الذي يتم تحديث مواعيده.
 * @param {Array<Object>} selectedSlots - مصفوفة من كائنات { timeSlotHeader: "HH:MM - HH:MM" } تمثل المواعيد التي يجب أن تكون متاحة.
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateTeacherSlots(teacherId, day, selectedSlots) {
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' غير موجود.");

    const data = teachersAvailableSlotsSheet.getDataRange().getValues();
    const headers = data[0]; // صف العناوين

    let teacherRowIndex = -1; // مؤشر الصف في مصفوفة البيانات (0-based)
    let headerToColIndexMap = new Map(); // map headers to their 0-based column index

    // البحث عن صف المعلم واليوم
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === String(teacherId).trim() && String(data[i][1] || '').trim() === String(day).trim()) {
        teacherRowIndex = i;
        break;
      }
    }

    if (teacherRowIndex === -1) {
      throw new Error(`لم يتم العثور على صف مواعيد للمعلم ID ${teacherId} في يوم ${day}.`);
    }

    // بناء خريطة لرؤوس الأعمدة ومؤشراتها
    const startColIndexForSlots = 2; // العمود C (مؤشر 2)
    for (let i = startColIndexForSlots; i < headers.length; i++) {
      headerToColIndexMap.set(String(headers[i] || '').trim(), i);
    }

    const updates = []; // لتجميع التحديثات وتطبيقها مرة واحدة
    const existingRowData = data[teacherRowIndex]; // البيانات الحالية للصف المعلم هذا اليوم

    // مقارنة المواعيد الموجودة مع المواعيد المختارة وتحديد التحديثات
    headerToColIndexMap.forEach((colIndex, timeSlotHeader) => {
      const currentCellValue = String(existingRowData[colIndex] || '').trim();
      const isSelectedInModal = selectedSlots.some(s => s.timeSlotHeader === timeSlotHeader);

      // إذا كانت الخلية محجوزة بالفعل (تحتوي على Student ID أو نص مخصص)، لا تغيرها
      // يجب أن تكون دالة `getTeacherAvailableSlots` هي من تُرسل ما إذا كان الموعد محجوزًا
      // هنا نفترض أننا لا نلمس الخلايا المحجوزة
      if (currentCellValue.startsWith("STD") || (currentCellValue !== '' && currentCellValue !== timeSlotHeader)) {
        // إذا كان الموعد محجوزاً، تأكد أنه ما زال محجوزاً، لا نلمس هذه الخلية
        // ولا تحاول إلغاء تحديدها حتى لو لم يتم اختيارها في المودال
        Logger.log(`الموعد ${day} ${timeSlotHeader} محجوز حالياً (${currentCellValue}). لن يتم تعديله.`);
        return; // انتقل إلى الموعد التالي
      }

      // الحالة: الموعد يجب أن يصبح متاحاً
      if (isSelectedInModal && currentCellValue !== timeSlotHeader) {
        updates.push({
          row: teacherRowIndex + 1,
          col: colIndex + 1,
          value: timeSlotHeader // اجعلها متاحة بوضع رأس العمود
        });
        Logger.log(`جعل الموعد ${day} ${timeSlotHeader} متاحاً.`);
      }
      // الحالة: الموعد يجب أن يصبح غير متاح (فارغ)
      else if (!isSelectedInModal && currentCellValue === timeSlotHeader) {
        updates.push({
          row: teacherRowIndex + 1,
          col: colIndex + 1,
          value: '' // اجعلها فارغة (غير متاحة)
        });
        Logger.log(`جعل الموعد ${day} ${timeSlotHeader} غير متاح.`);
      }
    });

    // تطبيق جميع التحديثات
    if (updates.length > 0) {
      const sheetBatchUpdates = [];
      updates.forEach(update => {
        sheetBatchUpdates.push([update.row, update.col, update.value]);
      });
      
      // لتطبيق التحديثات، ستحتاج إلى جلب النطاق المحدد وتحديثه.
      // الطريقة الأبسط هي التكرار على التحديثات الفردية، أو بناء مصفوفة كاملة للنطاق لتحديثه مرة واحدة.
      // للحفاظ على البساطة هنا، سنقوم بتطبيق التحديثات الفردية (قد تكون أبطأ مع عدد كبير جداً من التحديثات).
      updates.forEach(update => {
          teachersAvailableSlotsSheet.getRange(update.row, update.col).setValue(update.value);
      });
      Logger.log(`تم تطبيق ${updates.length} تحديث على مواعيد المعلم ${teacherId} في يوم ${day}.`);
    } else {
      Logger.log("لا توجد تغييرات لتطبيقها على مواعيد المعلم.");
    }

    return { success: "تم تحديث مواعيد المعلم بنجاح." };

  } catch (e) {
    Logger.log("خطأ في updateTeacherSlots: " + e.message);
    return { error: `فشل تحديث المواعيد: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}



// ==============================================================================
// 8. الدوال الخاصة بصفحة أرشيف الطلاب (Archive Page)
// ==============================================================================

/**
 * تجلب بيانات جميع الطلاب من شيت "أرشيف الطلاب".
 *
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب المؤرشفين.
 * أو {Object} كائن خطأ إذا كان شيت "أرشيف الطلاب" غير موجود.
 */
function getArchivedStudentsData() {
  const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("أرشيف الطلاب");

  if (!archiveSheet) {
    Logger.log("خطأ: شيت 'أرشيف الطلاب' غير موجود.");
    return { error: "شيت 'أرشيف الطلاب' غير موجود." };
  }

  const archivedData = archiveSheet.getDataRange().getValues();
  const archivedStudents = [];

  // ترتيب الأعمدة في شيت "أرشيف الطلاب":
  // Student ID(A), اسم الطالب(B), السن(C), رقم الهاتف (ولي الأمر)(D), رقم هاتف الطالب (إن وجد)(E), البلد(F),
  // تاريخ التسجيل(G), الحالة الأساسية للطالب(H), ملاحظات(I), تاريخ الأرشفة(J), سبب الأرشفة(K), تفاصيل الاشتراك وقت الأرشفة(L)

  archivedData.forEach((row, index) => {
    if (index === 0) return; // تخطي صف العناوين

    archivedStudents.push({
      rowIndex: index + 1, // رقم الصف في الأرشيف (1-based)
      studentID: String(row[0] || '').trim(),               // A
      name: String(row[1] || '').trim(),                    // B
      age: row[2],                                          // C
      phone: String(row[3] || '').trim(),                   // D (رقم هاتف ولي الأمر)
      studentPhone: String(row[4] || '').trim(),            // E (رقم هاتف الطالب)
      country: String(row[5] || '').trim(),                // F
      registrationDate: row[6] ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd") : '', // G
      basicStatus: String(row[7] || '').trim(),             // H (الحالة الأساسية للطالب وقت الأرشفة)
      notes: String(row[8] || '').trim(),                   // I
      archivedDate: row[9] ? Utilities.formatDate(row[9], Session.getScriptTimeZone(), "yyyy-MM-dd") : '', // J
      archiveReason: String(row[10] || '').trim(),          // K
      archiveSubscriptionDetails: String(row[11] || '').trim() // L
      // ملاحظة: لا يوجد renewalStatus أو packageName في الأرشيف مباشرة، يمكن إضافتهم كجزء من تفاصيل الاشتراك وقت الأرشفة إذا لزم الأمر
    });
  });

  Logger.log("تم جلب " + archivedStudents.length + " طالب من الأرشيف.");
  return archivedStudents;
}

/**
 * تعيد تفعيل طالب من الأرشيف إلى شيت "الطلاب".
 *
 * @param {string} studentID - معرف الطالب المراد إعادة تفعيله.
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function reactivateStudentFromArchive(studentID) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("أرشيف الطلاب");
  // المواعيد لا يتم إعادة حجزها تلقائياً عند إعادة التفعيل، يجب أن يتم ذلك يدوياً بعد التفعيل.

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!archiveSheet) throw new Error("شيت 'أرشيف الطلاب' غير موجود.");

    const archivedData = archiveSheet.getDataRange().getValues();
    let archiveRowIndex = -1; // الصف في الأرشيف (0-based)
    let studentToReactivateData = null;

    // البحث عن الطالب في الأرشيف
    for (let i = 1; i < archivedData.length; i++) {
      if (String(archivedData[i][0] || '').trim() === String(studentID).trim()) { // العمود A: Student ID
        archiveRowIndex = i;
        studentToReactivateData = archivedData[i];
        break;
      }
    }

    if (archiveRowIndex === -1 || !studentToReactivateData) {
      throw new Error(`لم يتم العثور على الطالب بمعرف ${studentID} في الأرشيف.`);
    }

    // نقل بيانات الطالب من الأرشيف إلى شيت "الطلاب"
    // أعمدة الأرشيف: Student ID(A), اسم الطالب(B), السن(C), رقم الهاتف (ولي الأمر)(D), رقم هاتف الطالب (إن وجد)(E), البلد(F), تاريخ التسجيل(G), الحالة الأساسية للطالب(H), ملاحظات(I), تاريخ الأرشفة(J), سبب الأرشفة(K), تفاصيل الاشتراك وقت الأرشفة(L)
    // أعمدة الطلاب: Student ID(A), اسم الطالب(B), السن(C), رقم الهاتف (ولي الأمر)(D), رقم هاتف الطالب (إن وجد)(E), البلد(F), تاريخ التسجيل(G), الحالة الأساسية للطالب(H), ملاحظات(I)
    
    const newStudentRow = [
      studentToReactivateData[0], // Student ID
      studentToReactivateData[1], // اسم الطالب
      studentToReactivateData[2], // السن
      studentToReactivateData[3], // رقم الهاتف (ولي الأمر)
      studentToReactivateData[4], // رقم هاتف الطالب (إن وجد)
      studentToReactivateData[5], // البلد
      studentToReactivateData[6], // تاريخ التسجيل الأصلي
      "معلق", // يمكن تعيين حالة مبدئية هنا، مثل "معلق" أو "قيد التسجيل" بعد إعادة التفعيل
      studentToReactivateData[8] // ملاحظات
    ];

    studentsSheet.appendRow(newStudentRow);
    Logger.log(`تمت إضافة الطالب ${studentToReactivateData[1]} (ID: ${studentID}) من الأرشيف إلى شيت 'الطلاب'.`);

    // حذف الطالب من شيت "أرشيف الطلاب"
    archiveSheet.deleteRow(archiveRowIndex + 1); // +1 لأن deleteRow تعمل بـ 1-based index

    Logger.log(`تمت إزالة الطالب ${studentToReactivateData[1]} (ID: ${studentID}) من الأرشيف.`);
    return { success: `تمت إعادة تفعيل الطالب ${studentToReactivateData[1]} بنجاح. يرجى مراجعة بياناته وتحديث اشتراكه ومواعيده.` };

  } catch (e) {
    Logger.log("خطأ في reactivateStudentFromArchive: " + e.message);
    return { error: `فشل إعادة تفعيل الطالب: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================================
// 9. الدوال الخاصة بالـ Triggers (المشغلات)
// ==============================================================================

/**
 * دالة التريجر التي يتم تشغيلها تلقائياً عند أي تعديل في Google Sheet.
 * (يجب ربطها بتريجر قابل للتثبيت "On edit").
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - كائن الحدث الذي يحتوي على معلومات حول التعديل.
 */
function handleOnEditTrigger(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedColumn = range.getColumn();
  const editedRow = range.getRow();
  const newValue = String(e.value || '').trim();
  const oldValue = String(e.oldValue || '').trim();

  // ========================================================================
  // منطق تصفير الحصص وتحديث تاريخ التجديد في شيت "الاشتراكات الحالية"
  // ========================================================================
  if (sheet.getName() === "الاشتراكات الحالية") {
    // إذا تم التعديل في عمود "الحالة التفصيلية للتجديد" (العمود H، مؤشر 7)
    if (editedColumn === 8) { // العمود H هو العمود الثامن
      if (editedRow > 1) { // التأكد من أنه ليس صف العناوين
        // إذا تم تغيير حالة التجديد من "يحتاج للتجديد" إلى "تم التجديد"
        if (oldValue === "يحتاج للتجديد" && newValue === "تم التجديد") {
          const studentId = String(sheet.getRange(editedRow, 2).getValue() || '').trim(); // العمود B: Student ID
          const packageName = String(sheet.getRange(editedRow, 3).getValue() || '').trim(); // العمود C: اسم الباقة

          // لا يتم تصفير الحصص للباقات نصف السنوية أو المخصصة (حيث لا تعتمد على عدد الحصص)
          const packageDetails = getPackageDetails(packageName); // دالة مساعدة
          const totalSessions = packageDetails ? packageDetails['عدد الحصص الكلي'] : 0;

          // إذا كانت باقة تعتمد على عدد الحصص (وليس نصف سنوية أو مخصصة لا يتم تصفيرها)
          if (totalSessions > 0 && packageName !== "اشتراك نصف سنوي" && packageName !== "مخصص") {
            // تصفير عدد الحصص الحاضرة في شيت "الاشتراكات الحالية"
            sheet.getRange(editedRow, 7).setValue(0); // العمود G: عدد الحصص الحاضرة
            Logger.log(`تم تصفير عدد الحصص الحاضرة للطالب ${studentId} في شيت 'الاشتراكات الحالية' إلى 0.`);
          } else {
            Logger.log(`اشتراك الطالب ${studentId} من نوع لا يتم تصفير حصصه (نصف سنوي/مخصص/غير محدد).`);
          }

          // تحديث تاريخ آخر تجديد إلى تاريخ اليوم
          const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
          sheet.getRange(editedRow, 9).setValue(today); // العمود I: تاريخ آخر تجديد
          Logger.log(`تم تحديث تاريخ آخر تجديد للطالب ${studentId} إلى ${today}.`);
        }
      }
    }
  }
}

/**
 * دالة لمعالجة بيانات الحضور وتحديث حالة التجديد للاشتراكات.
 * هذه الدالة يجب أن يتم تشغيلها بواسطة تريجر زمني (مثلاً كل 24 ساعة).
 */
function processAttendanceAndRenewal() {
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000); // انتظر حتى 60 ثانية

    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");
    if (!attendanceLogSheet) throw new Error("شيت 'سجل الحضور' غير موجود.");
    if (!packagesSheet) throw new Error("شيت 'الباقات' غير موجود.");

    const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
    const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
    const packagesData = packagesSheet.getDataRange().getValues();

    if (subscriptionsData.length < 2) {
      Logger.log("لا توجد بيانات اشتراكات لمعالجتها.");
      return;
    }

    // بناء خريطة لعدد الحصص لكل باقة
    const packageTotalSessionsMap = new Map();
    packagesData.forEach((row, index) => {
        if (index === 0) return;
        const packageName = String(row[0] || '').trim(); // العمود A: اسم الباقة
        const totalSessions = row[2]; // العمود C: عدد الحصص الكلي
        if (packageName && typeof totalSessions === 'number') {
            packageTotalSessionsMap.set(packageName, totalSessions);
        }
    });

    // حساب عدد الحصص الحاضرة لكل طالب من سجل الحضور (للحصص العادية فقط)
    const studentAttendedSessionsCount = new Map(); // key: Student ID, value: count
    attendanceLogData.forEach((row, index) => {
        if (index === 0) return; // تخطي العناوين
        const studentId = String(row[1] || '').trim(); // العمود B: Student ID
        const classType = String(row[8] || '').trim(); // العمود I: نوع الحصة

        if (studentId && classType === "عادية") { // فقط للحصص العادية
            studentAttendedSessionsCount.set(studentId, (studentAttendedSessionsCount.get(studentId) || 0) + 1);
        }
    });
    Logger.log("تم حساب الحصص الحاضرة من سجل الحضور.");


    const updatesForSubscriptionsSheet = []; // لتجميع التحديثات وتطبيقها مرة واحدة
    const today = new Date();
    today.setHours(0, 0, 0, 0); // لتجاهل الوقت عند المقارنة

    // المرور على الاشتراكات وتحديثها
    for (let i = 1; i < subscriptionsData.length; i++) {
        const row = subscriptionsData[i];
        const subscriptionRowIndex = i + 1; // الصف في الشيت (1-based)

        const studentId = String(row[1] || '').trim(); // العمود B: Student ID
        const packageName = String(row[2] || '').trim(); // العمود C: اسم الباقة
        const currentAttendedSessions = row[6] || 0; // العمود G: عدد الحصص الحاضرة
        const currentRenewalStatus = String(row[7] || '').trim(); // العمود H: الحالة التفصيلية للتجديد
        const lastRenewalDate = row[8]; // العمود I: تاريخ آخر تجديد


        // أ. تحديث عدد الحصص الحاضرة من سجل الحضور
        const newAttendedSessionsCount = studentAttendedSessionsCount.get(studentId) || 0;
        if (newAttendedSessionsCount !== currentAttendedSessions) {
            updatesForSubscriptionsSheet.push({
                rowIndex: subscriptionRowIndex,
                colIndex: 7, // العمود G: عدد الحصص الحاضرة
                value: newAttendedSessionsCount
            });
            Logger.log(`تحديث الحصص الحاضرة للطالب ${studentId}: من ${currentAttendedSessions} إلى ${newAttendedSessionsCount}`);
        }

        // ب. تحديث "الحالة التفصيلية للتجديد"
        let newRenewalStatus = currentRenewalStatus;
        const totalSessionsForPackage = packageTotalSessionsMap.get(packageName) || 0;

        // لا نغير حالة التجديد للباقات التجريبية أو نصف السنوية هنا
        if (packageName === "اشتراك نصف سنوي" || packageName === "مخصص" || currentRenewalStatus === "تجريبي") {
            // لا نغير حالتها بناءً على عدد الحصص
            // لكن نتحقق من انتهاء المدة للباقات نصف السنوية
            if (packageName === "اشتراك نصف سنوي" && lastRenewalDate instanceof Date) {
                 const sixMonthsLater = new Date(lastRenewalDate);
                 sixMonthsLater.setMonth(sixMonthsLater.getMonth() + 6);
                 sixMonthsLater.setHours(0, 0, 0, 0);
                 if (today >= sixMonthsLater && currentRenewalStatus !== "انتهت مدة الاشتراك") {
                     newRenewalStatus = "انتهت مدة الاشتراك";
                 }
            } else {
                // إذا كانت تجريبية أو مخصصة لا يتم تحديثها تلقائيًا بناءً على الحصص أو المدة
                // يمكن إضافة منطق لانتهاء صلاحية التجريبية هنا
            }
        } else if (totalSessionsForPackage > 0) { // باقات محددة الحصص
            if (newAttendedSessionsCount < totalSessionsForPackage) {
                newRenewalStatus = "تم التجديد"; // مازال لديه حصص
            } else if (newAttendedSessionsCount === totalSessionsForPackage) {
                newRenewalStatus = "يحتاج للتجديد"; // استهلك كل الحصص
            } else { // newAttendedSessionsCount > totalSessionsForPackage
                newRenewalStatus = "تجاوز الحد"; // تجاوز عدد الحصص
            }
        }
        
        // التحقق من انتهاء المدة (خاصة لو لم يتم تجديده يدوياً وكان "تم التجديد" ومازال لديه حصص)
        if (lastRenewalDate instanceof Date && (newRenewalStatus === "تم التجديد" || newRenewalStatus === "يحتاج للتجديد")) {
            const oneMonthLater = new Date(lastRenewalDate);
            oneMonthLater.setMonth(oneMonthLater.getMonth() + 1);
            oneMonthLater.setHours(0, 0, 0, 0);
            if (today >= oneMonthLater && newRenewalStatus !== "انتهت مدة الاشتراك" && newRenewalStatus !== "تجاوز الحد") {
                newRenewalStatus = "انتهت مدة الاشتراك";
            }
        }
        
        if (newRenewalStatus !== currentRenewalStatus) {
            updatesForSubscriptionsSheet.push({
                rowIndex: subscriptionRowIndex,
                colIndex: 8, // العمود H: الحالة التفصيلية للتجديد
                value: newRenewalStatus
            });
            Logger.log(`حالة تجديد الطالب ${studentId} تم تحديثها من ${currentRenewalStatus} إلى ${newRenewalStatus}.`);
        }
    }

    // تطبيق جميع التحديثات على شيت "الاشتراكات الحالية" مرة واحدة
    if (updatesForSubscriptionsSheet.length > 0) {
      updatesForSubscriptionsSheet.forEach(update => {
        subscriptionsSheet.getRange(update.rowIndex, update.colIndex).setValue(update.value);
      });
      Logger.log(`تم تطبيق ${updatesForSubscriptionsSheet.length} تحديث على شيت 'الاشتراكات الحالية'.`);
    }

    Logger.log("اكتملت عملية معالجة الحضور وتجديد الاشتراكات.");

  } catch (e) {
    Logger.log("خطأ عام في processAttendanceAndRenewal: " + e.message);
  } finally {
    lock.releaseLock();
  }
}



// ==============================================================================
// 4. الدوال الخاصة بصفحة تعديل الطلاب (Edit Student Page)
// ==============================================================================

/**
 * تجلب كافة البيانات الخاصة بطالب واحد بناءً على Student ID.
 * تجمع البيانات من شيتات "الطلاب"، "الاشتراكات الحالية"، و "المواعيد المتاحة للمعلمين".
 *
 * @param {string} studentId - Student ID للطالب المراد جلب بياناته.
 * @returns {Object|null} كائن الطالب الشامل أو null إذا لم يتم العثور عليه.
 */
function getStudentDataByID(studentId) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");

  if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");

  const studentData = studentsSheet.getDataRange().getValues();
  let studentFound = null; // الكائن الذي سيتم إرجاعه

  // 1. البحث عن الطالب في شيت "الطلاب"
  let studentRowIndex = -1;
  for (let i = 1; i < studentData.length; i++) {
    if (String(studentData[i][0] || '').trim() === String(studentId).trim()) { // العمود A: Student ID
      studentBasicInfo = {
        rowIndex: i + 1, // الصف في الشيت (1-based)
        studentID: String(studentData[i][0] || '').trim(),
        name: String(studentData[i][1] || '').trim(),
        age: studentData[i][2],
        phone: String(studentData[i][3] || '').trim(), // رقم الهاتف ولي الأمر
        studentPhone: String(studentData[i][4] || '').trim(), // رقم هاتف الطالب (إن وجد)
        country: String(studentData[i][5] || '').trim(),
        registrationDate: studentData[i][6] ? Utilities.formatDate(studentData[i][6], Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
        basicStatus: String(studentData[i][7] || '').trim(),
        notes: String(studentData[i][8] || '').trim()
      };
      studentFound = { ...studentBasicInfo }; // نسخ المعلومات الأساسية
      break;
    }
  }

  if (!studentFound) {
    Logger.log(`الطالب بمعرف ${studentId} غير موجود في شيت 'الطلاب'.`);
    return null;
  }

  // 2. دمج بيانات الاشتراك من شيت "الاشتراكات الحالية"
  if (subscriptionsSheet) {
    const subscriptionsValues = subscriptionsSheet.getDataRange().getValues();
    for (let i = 1; i < subscriptionsValues.length; i++) {
      if (String(subscriptionsValues[i][1] || '').trim() === studentId) { // العمود B: Student ID
        studentFound.subscriptionId = String(subscriptionsValues[i][0] || '').trim(); // Subscription ID (A)
        studentFound.packageName = String(subscriptionsValues[i][2] || '').trim(); // اسم الباقة (C)
        studentFound.teacherId = String(subscriptionsValues[i][3] || '').trim(); // Teacher ID (D)
        studentFound.subscriptionStartDate = subscriptionsValues[i][4] ? Utilities.formatDate(subscriptionsValues[i][4], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        studentFound.subscriptionEndDate = subscriptionsValues[i][5] ? Utilities.formatDate(subscriptionsValues[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        studentFound.attendedSessions = subscriptionsValues[i][6]; // عدد الحصص الحاضرة (G)
        studentFound.renewalStatus = String(subscriptionsValues[i][7] || '').trim(); // الحالة التفصيلية للتجديد (H)
        studentFound.lastRenewalDate = subscriptionsValues[i][8] ? Utilities.formatDate(subscriptionsValues[i][8], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        studentFound.totalSubscriptionAmount = subscriptionsValues[i][9];
        studentFound.paidAmount = subscriptionsValues[i][10];
        studentFound.remainingAmount = subscriptionsValues[i][11];
        break; // نفترض أن الطالب له اشتراك واحد فقط حالياً
      }
    }
  }

  // 3. دمج اسم المعلم (بدلاً من Teacher ID)
  if (studentFound.teacherId) {
      studentFound.teacherName = getTeacherNameById(studentFound.teacherId) || studentFound.teacherId;
  } else {
      studentFound.teacherName = '';
  }


  // 4. دمج المواعيد المحجوزة من شيت "المواعيد المتاحة للمعلمين"
  // نبحث عن المواعيد التي يكون فيها Student ID المحجوز به هو الطالب الحالي
  studentFound.bookedSlots = []; // لتخزين جميع المواعيد المحجوزة لهذا الطالب
  if (teachersAvailableSlotsSheet) {
    const availableSlotsValues = teachersAvailableSlotsSheet.getDataRange().getValues();
    const headers = availableSlotsValues[0];

    const timeSlotHeaders = [];
    const startColIndexForSlots = 2; // العمود C
    for (let i = startColIndexForSlots; i < headers.length; i++) {
        const header = String(headers[i] || '').trim();
        if (header) {
            timeSlotHeaders.push({ index: i, header: header });
        }
    }

    for (let i = 1; i < availableSlotsValues.length; i++) {
        const row = availableSlotsValues[i];
        const teacherIdInSlot = String(row[0] || '').trim(); // العمود A: Teacher ID
        const dayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم

        timeSlotHeaders.forEach(colInfo => {
            const slotValue = String(row[colInfo.index] || '').trim();
            const timeSlotHeader = colInfo.header;

            if (slotValue === studentId) { // لو الخلية تحتوي على Student ID لهذا الطالب
                studentFound.bookedSlots.push({
                    teacherId: teacherIdInSlot,
                    day: dayOfWeek,
                    timeSlotHeader: timeSlotHeader,
                    actualSlotValue: slotValue // القيمة الفعلية في الخلية (Student ID)
                });
            }
        });
    }
  }
  
  // لتبسيط العرض في الواجهة (اليوم الأول، الميعاد الأول)
  studentFound.day1 = studentFound.bookedSlots[0] ? studentFound.bookedSlots[0].day : '';
  studentFound.time1 = studentFound.bookedSlots[0] ? studentFound.bookedSlots[0].timeSlotHeader : '';
  studentFound.day2 = studentFound.bookedSlots[1] ? studentFound.bookedSlots[1].day : '';
  studentFound.time2 = studentFound.bookedSlots[1] ? studentFound.bookedSlots[1].timeSlotHeader : '';


  Logger.log(`تم جلب بيانات الطالب ${studentId}: ${JSON.stringify(studentFound)}`);
  return studentFound;
}

/**
 * تجلب بيانات الطلاب الذين يتطابقون مع رقم هاتف معين.
 * تجمع البيانات من شيت "الطلاب".
 *
 * @param {string} phone - رقم الهاتف للبحث عنه.
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب (معلومات أساسية فقط).
 */
function getStudentsByPhone(phone) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");

  if (!studentsSheet) {
    Logger.log("❌ خطأ: لم يتم العثور على شيت 'الطلاب' في getStudentsByPhone.");
    return { error: "لم يتم العثور على شيت 'الطلاب'." };
  }

  const data = studentsSheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    Logger.log("⚠️ تحذير: لا توجد بيانات في شيت 'الطلاب'.");
    return [];
  }

  const students = [];
  // تنظيف رقم الهاتف للبحث (إزالة الصفر الأول إذا كان موجودًا)
  const searchPhone = String(phone).trim(); 
  const cleanedSearchPhone = searchPhone.startsWith("0") ? searchPhone.substring(1) : searchPhone;

  Logger.log("📞 البحث عن رقم الهاتف: " + cleanedSearchPhone);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const studentID = String(row[0] || '').trim(); // العمود A: Student ID
    const studentName = String(row[1] || '').trim(); // العمود B: اسم الطالب
    const studentAge = row[2]; // العمود C: السن
    const storedPhone = String(row[3] || '').trim(); // العمود D: رقم الهاتف (ولي الأمر)

    // تنظيف رقم الهاتف المخزن بنفس الطريقة للمقارنة
    const cleanedStoredPhone = storedPhone.startsWith("0") ? storedPhone.substring(1) : storedPhone;

    if (cleanedStoredPhone === cleanedSearchPhone) {
      Logger.log("Match found at row " + (i + 1));
      students.push({
        rowIndex: i + 1, // الصف في الشيت (1-based)
        studentID: studentID,
        name: studentName,
        age: studentAge,
        phone: storedPhone, // نرجع الرقم الأصلي من الشيت
        basicStatus: String(row[7] || '').trim() // الحالة الأساسية للطالب (H)
      });
    }
  }

  Logger.log("✅ عدد الطلاب الذين تم العثور عليهم برقم الهاتف: " + students.length);
  return students; // نرجع مصفوفة حتى لو طالب واحد، ليتم التعامل معها في الواجهة الأمامية
}


/**
 * تحديث بيانات الطالب في شيت "الطلاب" و "الاشتراكات الحالية"، وإعادة تخصيص المواعيد.
 *
 * @param {Object} updatedData - كائن يحتوي على البيانات المحدثة للطالب،
 * بالإضافة إلى oldTeacherName, oldDay1, oldTime1, oldDay2, oldTime2 (لتحرير المواعيد القديمة).
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateStudentDataWithReassignment(updatedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName("الطلاب");
  const subscriptionsSheet = ss.getSheetByName("الاشتراكات الحالية");
  const slotsSheet = ss.getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = ss.getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    if (!studentsSheet || !subscriptionsSheet || !slotsSheet || !teachersSheet) {
      throw new Error("تعذر العثور على أحد الشيتات المطلوبة.");
    }

    const studentId = updatedData.studentID;
    const studentRowIndex = updatedData.rowIndex;
    if (!studentId || !studentRowIndex) {
      throw new Error("الرقم التعريفي للطالب أو رقم الصف غير متوفر.");
    }

    // تحديث بيانات الطالب في شيت الطلاب
    studentsSheet.getRange(studentRowIndex, 2).setValue(updatedData.editName); // الاسم
    studentsSheet.getRange(studentRowIndex, 3).setValue(updatedData.editAge);  // السن
    studentsSheet.getRange(studentRowIndex, 4).setValue(updatedData.editPhone); // الهاتف

    // الحصول على Teacher ID
    const teacherId = getTeacherIdByName(updatedData.editTeacher);
    if (!teacherId) throw new Error(`لم يتم العثور على Teacher ID للمعلم ${updatedData.editTeacher}`);

    // جلب الاشتراك الحالي لهذا الطالب (نستخدم أول اشتراك غير تجريبي بشكل مؤقت)
    const allSubscriptions = subscriptionsSheet.getDataRange().getValues();
    const headers = allSubscriptions[0];
    const studentCol = headers.indexOf("Student ID");
    const packageCol = headers.indexOf("اسم الباقة");
    const subscriptionIdCol = headers.indexOf("Subscription ID");
    const teacherIdCol = headers.indexOf("Teacher ID");
    const renewalStatusCol = headers.indexOf("الحالة التفصيلية للتجديد");

    let subRowIndex = -1;
    let subscriptionId = "";

    for (let i = 1; i < allSubscriptions.length; i++) {
      if (String(allSubscriptions[i][studentCol]).trim() === studentId) {
        subscriptionId = String(allSubscriptions[i][subscriptionIdCol]).trim();
        subRowIndex = i + 1; // 1-based
        break;
      }
    }

    if (!subscriptionId) throw new Error(`لم يتم العثور على اشتراك لهذا الطالب ${studentId}`);

    // تحديث بيانات الاشتراك
    subscriptionsSheet.getRange(subRowIndex, packageCol + 1).setValue(updatedData.editSubscriptionPackage);
    subscriptionsSheet.getRange(subRowIndex, teacherIdCol + 1).setValue(teacherId);
    subscriptionsSheet.getRange(subRowIndex, renewalStatusCol + 1).setValue(updatedData.editPaymentStatus);

    // تحرير المواعيد القديمة الخاصة بالـ Subscription ID
    const oldSlots = getBookedSlotsBySubscriptionId(slotsSheet, subscriptionId);
    oldSlots.forEach(slot => {
      releaseTeacherSlot(slotsSheet, slot.teacherId, slot.day, slot.timeSlotHeader, subscriptionId);
    });

    // إضافة المواعيد الجديدة
    const newSlots = [];
    if (updatedData.editDay1 && updatedData.editTime1) {
      newSlots.push({ day: updatedData.editDay1, time: updatedData.editTime1 });
    }
    if (updatedData.editDay2 && updatedData.editTime2) {
      newSlots.push({ day: updatedData.editDay2, time: updatedData.editTime2 });
    }

    const bookingType = updatedData.editPaymentStatus === "تجريبي" ? "تجريبي" : "عادي";

    newSlots.forEach(slot => {
      const result = bookTeacherSlot(slotsSheet, teacherId, slot.day, slot.time, subscriptionId, bookingType);
      if (result.error) throw new Error(`تعذر حجز الميعاد: ${slot.day} ${slot.time} - ${result.error}`);
    });

    Logger.log(`تم تحديث بيانات الطالب ${studentId} والمواعيد بنجاح.`);
    return { success: "تم حفظ التعديلات بنجاح." };

  } catch (e) {
    Logger.log("خطأ في updateStudentDataWithReassignment: " + e.message);
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}



/**
 * تسترجع كل المواعيد المحجوزة المرتبطة بـ Subscription ID معيّن من شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {Sheet} slotsSheet - شيت "المواعيد المتاحة للمعلمين".
 * @param {string} subscriptionId - معرف الاشتراك المراد البحث عنه.
 * @returns {Array<Object>} قائمة المواعيد المحجوزة (teacherId, day, timeSlotHeader, bookingType).
 */
function getBookedSlotsBySubscriptionId(slotsSheet, subscriptionId) {
  const allData = slotsSheet.getDataRange().getValues();
  const headers = allData[0];
  const results = [];

  for (let rowIndex = 1; rowIndex < allData.length; rowIndex++) {
    const row = allData[rowIndex];
    const teacherId = String(row[0] || '').trim(); // العمود A
    const day = String(row[1] || '').trim();       // العمود B

    for (let colIndex = 2; colIndex < row.length; colIndex++) {
      const cellValue = String(row[colIndex] || '').trim();
      
      if (cellValue) {
        // دعم صيغة: SUB1234 - نور بيان أو SUB5678 - عادي
        const [cellSubscriptionId, type] = cellValue.split(' - ').map(s => s.trim());

        if (cellSubscriptionId === subscriptionId) {
          const timeSlotHeader = headers[colIndex];
          const bookingType = type || "عادي"; // افتراضي لو مش محدد
          results.push({
            teacherId: teacherId,
            day: day,
            timeSlotHeader: timeSlotHeader,
            bookingType: bookingType
          });
        }
      }
    }
  }

  return results;
}



/**
 * دالة مساعدة: تجلب تفاصيل اشتراك طالب واحد من شيت "الاشتراكات الحالية".
 * @param {string} studentId - Student ID للطالب.
 * @returns {Object|null} كائن الاشتراك أو null.
 */
function getStudentSubscriptionByStudentId(studentId) {
    const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
    if (!subscriptionsSheet) return null;

    const data = subscriptionsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][1] || '').trim() === String(studentId).trim()) { // العمود B: Student ID
            return {
                rowIndex: i + 1, // الصف في الشيت (1-based)
                subscriptionId: String(data[i][0] || '').trim(), // العمود A
                subscriptionType: String(data[i][2] || '').trim(), // ✅ العمود C (نوع الاشتراك)
                packageName: String(data[i][3] || '').trim(),      // ✅ العمود D (اسم الباقة)
                teacherId: String(data[i][4] || '').trim(),        // ✅ العمود E (Teacher ID)
                // ... يمكن إضافة باقي الأعمدة حسب الحاجة
            };
        }
    }
    return null;
}


/**
 * دالة مساعدة: تجلب جميع المواعيد المحجوزة بواسطة طالب معين من شيت "المواعيد المتاحة للمعلمين".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - شيت "المواعيد المتاحة للمعلمين".
 * @param {string} studentId - Student ID للطالب.
 * @returns {Array<Object>} مصفوفة من المواعيد المحجوزة.
 */
function getBookedSlotsByStudentId(sheet, studentId) {
    const bookedSlots = [];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const timeSlotColIndexes = [];
    const startColIndexForSlots = 2; // العمود C
    for (let i = startColIndexForSlots; i < headers.length; i++) {
        const header = String(headers[i] || '').trim();
        if (header) {
            timeSlotColIndexes.push(i);
        }
    }

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const teacherId = String(row[0] || '').trim(); // العمود A: Teacher ID
        const dayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم

        timeSlotColIndexes.forEach(colIndex => {
            const slotValue = String(row[colIndex] || '').trim();
            const timeSlotHeader = headers[colIndex];

            // إذا كانت الخلية تحتوي على Student ID لهذا الطالب
            if (slotValue === studentId) {
                bookedSlots.push({
                    teacherId: teacherId,
                    day: dayOfWeek,
                    timeSlotHeader: timeSlotHeader,
                    rowIndex: i + 1, // رقم الصف في شيت المواعيد المتاحة (1-based)
                    colIndex: colIndex + 1 // رقم العمود في شيت المواعيد المتاحة (1-based)
                });
            }
        });
    }
    return bookedSlots;
}

/**
 * دالة مساعدة: تحرر موعداً (تجعل الخلية فارغة) في شيت "المواعيد المتاحة للمعلمين".
 * تستخدم في تحرير المواعيد القديمة عند التعديل أو الحذف.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - شيت "المواعيد المتاحة للمعلمين".
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} day - اليوم.
 * @param {string} timeSlotHeader - رأس الميعاد.
 * @param {string} expectedStudentId - الـ Student ID الذي يجب أن يكون في الخلية (للتأكيد).
 * @returns {Object} رسالة نجاح أو خطأ.
 */
function releaseTeacherSlot(sheet, teacherId, day, timeSlotHeader, expectedStudentId) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    let teacherRowIndex = -1;
    let timeSlotColIndex = -1;

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === String(teacherId).trim() && String(data[i][1] || '').trim() === String(day).trim()) {
            teacherRowIndex = i;
            break;
        }
    }
    if (teacherRowIndex === -1) return { error: `لم يتم العثور على صف المعلم/اليوم لتحرير الميعاد: ${teacherId}, ${day}` };

    for (let i = 2; i < headers.length; i++) {
        if (String(headers[i] || '').trim() === String(timeSlotHeader).trim()) {
            timeSlotColIndex = i;
            break;
        }
    }
    if (timeSlotColIndex === -1) return { error: `لم يتم العثور على عمود الميعاد لتحريره: ${timeSlotHeader}` };

    const targetCell = sheet.getRange(teacherRowIndex + 1, timeSlotColIndex + 1);
    const currentCellValue = String(targetCell.getValue() || '').trim();

    // تأكيد أن الميعاد محجوز بالفعل بواسطة هذا الطالب قبل تحريره
    if (currentCellValue === expectedStudentId) {
        targetCell.setValue(timeSlotHeader);
        Logger.log(`تم تحرير الميعاد ${day} ${timeSlotHeader} للمعلم ${teacherId} من الطالب ${expectedStudentId}.`);
        return { success: 'تم تحرير الميعاد بنجاح.' };
    } else {
        Logger.log(`الميعاد ${day} ${timeSlotHeader} للمعلم ${teacherId} ليس محجوزاً بواسطة الطالب ${expectedStudentId} (قيمته: ${currentCellValue}).`);
        return { error: 'الميعاد ليس محجوزاً بهذا الطالب أو محجوز بنص آخر.' };
    }
}


/**
 * حذف طالب ونقله إلى الأرشيف.
 * @param {Object} studentInfo - كائن يحتوي على studentID ورقم الصف في شيت الطلاب.
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function deleteStudentAndArchive(studentInfo) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  let archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("أرشيف الطلاب");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");
    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' غير موجود.");
    if (!archiveSheet) {
      archiveSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("أرشيف الطلاب");
      archiveSheet.appendRow([
        "Student ID", "اسم الطالب", "السن", "رقم الهاتف (ولي الأمر)", "رقم هاتف الطالب (إن وجد)", "البلد",
        "تاريخ التسجيل", "الحالة الأساسية للطالب", "ملاحظات",
        "تاريخ الأرشفة", "سبب الأرشفة", "تفاصيل الاشتراك وقت الأرشفة" // أعمدة إضافية للأرشيف
      ]);
    }

    const studentIdToDelete = studentInfo.studentID;
    const studentRowInStudentsSheet = studentInfo.rowIndex;

    // 1. جلب بيانات الطالب الكاملة من شيت "الطلاب" قبل الحذف
    const studentDataToArchive = studentsSheet.getRange(studentRowInStudentsSheet, 1, 1, 9).getValues()[0]; // من A إلى I
    if (!studentDataToArchive || studentDataToArchive.length === 0) {
      throw new Error(`لم يتم العثور على بيانات الطالب ID ${studentIdToDelete} في الصف المحدد للحذف.`);
    }

    // 2. تحرير جميع المواعيد التي حجزها هذا الطالب في شيت "المواعيد المتاحة للمعلمين"
    const bookedSlotsToRelease = getBookedSlotsByStudentId(teachersAvailableSlotsSheet, studentIdToDelete);
    bookedSlotsToRelease.forEach(slot => {
        releaseTeacherSlot(teachersAvailableSlotsSheet, slot.teacherId, slot.day, slot.timeSlotHeader, studentIdToDelete);
    });
    Logger.log(`تم تحرير ${bookedSlotsToRelease.length} موعد للطالب ${studentIdToDelete}.`);


    // 3. نقل بيانات الطالب إلى شيت "أرشيف الطلاب"
    const archivedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const reason = studentInfo.name ? `حذف من النظام (${studentInfo.name})` : "حذف من النظام"; // يمكن تخصيص السبب أكثر
    
    // جلب تفاصيل الاشتراك الأخيرة لإضافتها للأرشيف
    let lastSubscriptionDetails = "لا توجد تفاصيل اشتراك.";
    const studentSubscription = getStudentSubscriptionByStudentId(studentIdToDelete);
    if (studentSubscription) {
        lastSubscriptionDetails = `Subscription ID: ${studentSubscription.subscriptionId}, Package: ${studentSubscription.packageName}, Teacher ID: ${studentSubscription.teacherId}`;
        // يمكنك توسيع هذا النص ليشمل المزيد من التفاصيل
    }

    archiveSheet.appendRow([
      ...studentDataToArchive, // كل الأعمدة الأصلية من شيت الطلاب
      archivedDate,           // تاريخ الأرشفة (J)
      reason,                 // سبب الأرشفة (K)
      lastSubscriptionDetails // تفاصيل الاشتراك وقت الأرشفة (L)
    ]);
    Logger.log(`تم أرشفة الطالب ${studentDataToArchive[1]} (ID: ${studentIdToDelete}).`);


    // 4. حذف صف الطالب من شيت "الطلاب"
    studentsSheet.deleteRow(studentRowInStudentsSheet);
    Logger.log(`تم حذف صف الطالب ${studentIdToDelete} من شيت 'الطلاب'.`);

    // 5. حذف اشتراك الطالب من شيت "الاشتراكات الحالية"
    if (studentSubscription) { // تأكد من وجود اشتراك
        subscriptionsSheet.deleteRow(studentSubscription.rowIndex);
        Logger.log(`تم حذف اشتراك الطالب ${studentIdToDelete} من شيت 'الاشتراكات الحالية'.`);
    }

    return { success: `تم حذف الطالب ${studentDataToArchive[1]} ونقله إلى الأرشيف بنجاح.` };

  } catch (e) {
    Logger.log("خطأ في deleteStudentAndArchive: " + e.message);
    return { error: `فشل حذف الطالب: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}



// ==============================================================================
// 10. الدوال الخاصة بصفحة تسجيل الطلاب التجريبيين (Trial Registration Page)
// ==============================================================================

/**
 * حفظ بيانات الطالب التجريبي الجديد.
 * تقوم بحفظ البيانات في شيت "الطلاب التجريبيون"،
 * وتحديث الميعاد المحجوز في شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {Object} formData - كائن يحتوي على بيانات النموذج المرسلة من الواجهة الأمامية.
 * المتوقع من الواجهة الأمامية (من حقول تسجيل الطالب التجريبي):
 * { trialRegName, trialRegAge, trialRegPhone, trialRegTeacher (اسم المعلم), trialRegDay1, trialRegTime1 }
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function saveTrialData(formData) {
  const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!trialStudentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب التجريبيون'.");
    if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    if (!teachersSheet) throw new Error("لم يتم العثور على شيت 'المعلمين'.");

    // 1. جلب Teacher ID من اسم المعلم
    const teacherId = getTeacherIdByName(formData.trialRegTeacher);
    if (!teacherId) throw new Error(`لم يتم العثور على Teacher ID للمعلم: ${formData.trialRegTeacher}`);
    const teacherName = getTeacherNameById(teacherId);

    // 2. إنشاء Trial ID
    const newTrialId = generateUniqueTrialId(trialStudentsSheet);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // 3. إضافة الطالب في شيت "الطلاب التجريبيون"
    trialStudentsSheet.appendRow([
      newTrialId,             // A: Trial ID
      formData.trialRegName,  // B: الاسم
      formData.trialRegAge,   // C: السن
      String(formData.trialRegPhone).trim(), // D: الهاتف
      teacherId,              // E
      teacherName,            // F
      formData.trialRegDay1,  // G
      formData.trialRegTime1, // H
      today,                  // I: تاريخ التسجيل
      "",                     // J: ملاحظات
      "تجريبي"               // K: الحالة
    ]);
    Logger.log(`تم حفظ الطالب التجريبي ${formData.trialRegName} (${newTrialId}) في 'الطلاب التجريبيون'.`);

    // 4. إضافة اشتراك تجريبي في شيت "الاشتراكات الحالية"
    const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet);

    subscriptionsSheet.appendRow([
      newSubscriptionId,     // A: Subscription ID
      newTrialId,            // B: Student ID (نستخدم Trial ID)
      "تجريبي",              // ✅ C: نوع الاشتراك
      "",                    // D: اسم الباقة (فارغ لأنه تجريبي)
      teacherId,             // E
      today,                 // F: تاريخ البداية
      "",                    // G: تاريخ نهاية الاشتراك المتوقع
      0,                     // H: عدد الحصص الحاضرة
      "تجريبي",              // I: حالة التجديد
      today,                 // J: تاريخ آخر تجديد
      0,                     // K: السعر الكلي
      0,                     // L: المدفوع
      0,                     // M: المتبقي
      "طالب تجريبي"          // N: ملاحظات
    ]);
    Logger.log(`تم إنشاء اشتراك تجريبي (${newSubscriptionId}) للطالب ${newTrialId}.`);

    // 5. حجز الميعاد
    const result = bookTeacherSlot(
      teachersAvailableSlotsSheet,
      teacherId,
      formData.trialRegDay1,
      formData.trialRegTime1,
      newTrialId,
      "تجريبي"
    );
    if (result.error) {
      throw new Error(`تعذر حجز الميعاد: ${result.error}`);
    }

    Logger.log("اكتملت عملية حفظ الطالب التجريبي بنجاح.");
    return { success: "تم تسجيل الطالب التجريبي بنجاح." };

  } catch (e) {
    Logger.log("خطأ في saveTrialData: " + e.message);
    return { error: `فشل تسجيل الطالب التجريبي: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


// ==============================================================================
// 11. الدوال الخاصة بتحويل الطلاب التجريبيين إلى مشتركين
// ==============================================================================

/**
 * تحويل طالب تجريبي إلى طالب مشترك.
 * تنقل بيانات الطالب من شيت "الطلاب التجريبيون" إلى شيت "الطلاب"،
 * وتنشئ اشتراكًا جديدًا في شيت "الاشتراكات الحالية"،
 * وتحدث المواعيد المحجوزة في شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {string} trialStudentId - معرف الطالب التجريبي (TRLxxx).
 * @param {string} selectedPackageName - اسم الباقة التي اختارها الطالب (مثلاً: "باقة 8 حصص شهري").
 * @param {string} paymentStatus - حالة الدفع الأولية للاشتراك الجديد (مثلاً: "تم الدفع", "لم يتم الدفع").
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function convertTrialToRegistered(trialStudentId, selectedPackageName, paymentStatus) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // التحقق من وجود الشيتات الضرورية
    if (!studentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب'.");
    if (!trialStudentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب التجريبيون'.");
    if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");

    // 1. جلب بيانات الطالب التجريبي
    const trialStudentsData = trialStudentsSheet.getDataRange().getValues();
    let trialStudentRowIndex = -1;
    let trialStudentData = null;

    for (let i = 1; i < trialStudentsData.length; i++) {
      if (String(trialStudentsData[i][0] || '').trim() === String(trialStudentId).trim()) { // العمود A: Trial ID
        trialStudentRowIndex = i; // الصف في مصفوفة البيانات (0-based)
        trialStudentData = trialStudentsData[i];
        break;
      }
    }

    if (trialStudentRowIndex === -1 || !trialStudentData) {
      throw new Error(`لم يتم العثور على الطالب التجريبي بمعرف ${trialStudentId}.`);
    }

    // استخراج البيانات الأساسية للطالب التجريبي
    const name = String(trialStudentData[1] || '').trim(); // العمود B: Student Name
    const age = trialStudentData[2];                       // العمود C: Age
    const phone = String(trialStudentData[3] || '').trim(); // العمود D: Phone Number
    const teacherId = String(trialStudentData[4] || '').trim(); // العمود E: Teacher ID
    const day1 = String(trialStudentData[6] || '').trim(); // العمود G: Day
    const time1 = String(trialStudentData[7] || '').trim(); // العمود H: Time (الميعاد الأول)

    // 2. حفظ الطالب في شيت "الطلاب" (كطالب مشترك الآن)
    const newStudentId = generateUniqueStudentId(studentsSheet); // دالة موجودة
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // تحديد الحالة الأساسية للطالب (ستكون "مشترك" بعد التحويل)
    let studentBasicStatus = "مشترك";

    // ترتيب الأعمدة في شيت "الطلاب": Student ID, اسم الطالب, السن, رقم الهاتف (ولي الأمر), رقم هاتف الطالب (إن وجد), البلد, تاريخ التسجيل, الحالة الأساسية للطالب, ملاحظات
    studentsSheet.appendRow([
      newStudentId,                  // Student ID (A)
      name,                          // اسم الطالب (B)
      age,                           // السن (C)
      phone,                         // رقم الهاتف (ولي الأمر) (D)
      "",                            // رقم هاتف الطالب (إن وجد) (E) - يمكن إضافته لاحقًا أو من بيانات الطالب التجريبي
      "",                            // البلد (F)
      today,                         // تاريخ التسجيل (G)
      studentBasicStatus,            // الحالة الأساسية للطالب (H)
      "تم التحويل من تجريبي"         // ملاحظات (I)
    ]);
    Logger.log(`تم حفظ الطالب ${name} (ID: ${newStudentId}) في شيت 'الطلاب'.`);

    // 3. إنشاء اشتراك في شيت "الاشتراكات الحالية"
    const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet); // دالة موجودة
    const packageDetails = getPackageDetails(selectedPackageName); // دالة موجودة

    let subscriptionRenewalStatus = "تم التجديد"; // افتراضي للاشتراك الجديد
    let totalClassesAttended = 0; // يبدأ بـ 0 حصة حاضرة
    let subscriptionType = "شهري"; // افتراضي

    if (packageDetails && packageDetails['نوع الباقة']) {
      subscriptionType = packageDetails['نوع الباقة'];
    }

    // حساب تاريخ نهاية الاشتراك المتوقع
    let endDate = "";
    if (subscriptionType === "شهري") {
        const startDate = new Date(today);
        startDate.setMonth(startDate.getMonth() + 1);
        endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (subscriptionType === "نصف سنوي") {
        const startDate = new Date(today);
        startDate.setMonth(startDate.getMonth() + 6);
        endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    // تحديد المبلغ المدفوع والمتبقي بناءً على paymentStatus
    let paidAmount = 0;
    let remainingAmount = packageDetails ? packageDetails['السعر'] : 0;

    if (paymentStatus === "تم الدفع") {
        paidAmount = packageDetails ? packageDetails['السعر'] : 0;
        remainingAmount = 0;
    } else if (paymentStatus === "تم دفع جزء") {
        // ستحتاج لآلية لإدخال المبلغ المدفوع جزئياً
        // حالياً، سنفترض أنه لم يدفع شيئاً إذا لم يكن "تم الدفع" كاملاً
        paidAmount = 0; // أو يمكن أن يُمرر كمُدخل جديد للدالة
        remainingAmount = packageDetails ? packageDetails['السعر'] : 0;
    }

    // ترتيب الأعمدة في شيت "الاشتراكات الحالية": Subscription ID, Student ID, اسم الباقة, Teacher ID, تاريخ بداية الاشتراك, تاريخ نهاية الاشتراك المتوقع, عدد الحصص الحاضرة, الحالة التفصيلية للتجديد, تاريخ آخر تجديد, مبلغ الاشتراك الكلي, المبلغ المدفوع حتى الآن, المبلغ المتبقي, ملاحظات خاصة بالاشتراك
    subscriptionsSheet.appendRow([
      newSubscriptionId,             // Subscription ID (A)
      newStudentId,                  // Student ID (B)
      selectedPackageName,           // اسم الباقة (C)
      teacherId,                     // Teacher ID (D)
      today,                         // تاريخ بداية الاشتراك (E)
      endDate,                       // تاريخ نهاية الاشتراك المتوقع (F)
      totalClassesAttended,          // عدد الحصص الحاضرة (G)
      subscriptionRenewalStatus,     // الحالة التفصيلية للتجديد (H)
      today,                         // تاريخ آخر تجديد (I)
      packageDetails ? packageDetails['السعر'] : 0, // مبلغ الاشتراك الكلي (J)
      paidAmount,                    // المبلغ المدفوع حتى الآن (K)
      remainingAmount,               // المبلغ المتبقي (L)
      "تحويل من طالب تجريبي"         // ملاحظات خاصة بالاشتراك (M)
    ]);
    Logger.log(`تم إنشاء اشتراك (${newSubscriptionId}) للطالب ${newStudentId} في شيت 'الاشتراكات الحالية'.`);

    // 4. تحديث المواعيد في شيت "المواعيد المتاحة للمعلمين"
    // أ. تحرير الميعاد القديم المحجوز بمعرف الطالب التجريبي
    const releaseResult = releaseTeacherSlot(
        teachersAvailableSlotsSheet,
        teacherId,
        day1,
        time1,
        trialStudentId // تأكيد أن الميعاد محجوز بمعرف الطالب التجريبي
    );
    if (releaseResult.error) {
        Logger.log(`تحذير: خطأ في تحرير الميعاد القديم للطالب التجريبي ${trialStudentId}: ${releaseResult.error}`);
        // لا نرمي خطأ هنا لأنه ليس حرجًا للعملية ككل
    } else {
        Logger.log(`تم تحرير الميعاد القديم للطالب التجريبي ${trialStudentId}.`);
    }

    // ب. حجز الميعاد الجديد بمعرف الطالب المشترك
    const bookResult = bookTeacherSlot( // إعادة استخدام دالة bookTeacherSlot
        teachersAvailableSlotsSheet,
        teacherId,
        day1,
        time1,
        newStudentId, // استخدم Student ID الجديد
        "عادي" // نوع الحجز أصبح عادي
    );
    if (bookResult.error) {
        Logger.log(`خطأ في حجز الميعاد الجديد للطالب المشترك ${newStudentId}: ${bookResult.error}`);
        throw new Error(`تعذر حجز الميعاد للطالب المشترك: ${bookResult.error}`);
    } else {
        Logger.log(`تم حجز الميعاد الجديد للطالب المشترك ${newStudentId}.`);
    }

    // 5. حذف الطالب من شيت "الطلاب التجريبيون"
    trialStudentsSheet.deleteRow(trialStudentRowIndex + 1); // +1 لأنه 1-based index
    Logger.log(`تم حذف الطالب التجريبي ${trialStudentId} من شيت 'الطلاب التجريبيون'.`);

    return { success: `تم تحويل الطالب ${name} بنجاح إلى مشترك!` };

  } catch (e) {
    Logger.log("خطأ في convertTrialToRegistered: " + e.message);
    return { error: `فشل التحويل: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}




/**
 * تجلب كافة البيانات الخاصة بطالب تجريبي واحد بناءً على Trial ID.
 * تجمع البيانات من شيت "الطلاب التجريبيون"، والمعلمين، والمواعيد المحجوزة.
 *
 * @param {string} trialId - Trial ID للطالب التجريبي المراد جلب بياناته.
 * @returns {Object|null} كائن الطالب التجريبي الشامل أو null إذا لم يتم العثور عليه.
 */
function getTrialStudentDataByID(trialId) {
    const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");
    const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
    const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");

    if (!trialStudentsSheet) throw new Error("شيت 'الطلاب التجريبيون' غير موجود.");

    const trialStudentData = trialStudentsSheet.getDataRange().getValues();
    let studentFound = null;

    // 1. البحث عن الطالب في شيت "الطلاب التجريبيون"
    let studentRowIndex = -1;
    for (let i = 1; i < trialStudentData.length; i++) {
        if (String(trialStudentData[i][0] || '').trim() === String(trialId).trim()) { // العمود A: Trial ID
            studentFound = {
                rowIndex: i + 1, // الصف في الشيت (1-based)
                trialID: String(trialStudentData[i][0] || '').trim(), // A
                name: String(trialStudentData[i][1] || '').trim(),  // B
                age: trialStudentData[i][2],                         // C
                phone: String(trialStudentData[i][3] || '').trim(),  // D
                teacherId: String(trialStudentData[i][4] || '').trim(), // E
                // teacherName: String(trialStudentData[i][5] || '').trim(), // F
                day: String(trialStudentData[i][6] || '').trim(),   // G
                time: String(trialStudentData[i][7] || '').trim(),   // H
                registrationDate: trialStudentData[i][8] ? Utilities.formatDate(trialStudentData[i][8], Session.getScriptTimeZone(), "yyyy-MM-dd") : '', // I
                notes: String(trialStudentData[i][9] || '').trim(),  // J
                status: String(trialStudentData[i][10] || '').trim() // K
            };
            break;
        }
    }

    if (!studentFound) {
  Logger.log(`الطالب التجريبي بمعرف ${trialId} غير موجود.`);
  return { found: false, error: "لم يتم العثور على الطالب التجريبي بهذا المعرف." }; // <--- التعديل
}

    // 2. دمج اسم المعلم (من Teacher ID)
    if (studentFound.teacherId) {
        studentFound.teacherName = getTeacherNameById(studentFound.teacherId) || studentFound.teacherId;
    } else {
        studentFound.teacherName = '';
    }
    // (المواعيد الحالية للطالب التجريبي ستكون Day و Time مباشرة من شيت الطلاب التجريبيين)
    // لا نحتاج للبحث في teachersAvailableSlotsSheet لجلب مواعيده، لأنها مخزنة مباشرة في صفه.

    Logger.log(`تم جلب بيانات الطالب التجريبي ${trialId}: ${JSON.stringify(studentFound)}`);
    return studentFound;
}

/**
 * تجلب بيانات الطلاب التجريبيين الذين يتطابقون مع رقم هاتف معين.
 *
 * @param {string} phone - رقم الهاتف للبحث عنه.
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب التجريبيين.
 */
function getTrialStudentsByPhone(phone) {
    const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");
    const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين"); // لجلب اسم المعلم

    if (!trialStudentsSheet) {
        Logger.log("❌ خطأ: لم يتم العثور على شيت 'الطلاب التجريبيون' في getTrialStudentsByPhone.");
        return { error: "لم يتم العثور على شيت 'الطلاب التجريبيون'." };
    }

    const data = trialStudentsSheet.getDataRange().getValues();
    if (!data || data.length === 0) {
        Logger.log("⚠️ تحذير: لا توجد بيانات في شيت 'الطلاب التجريبيون'.");
        return [];
    }

    const students = [];
    const searchPhone = String(phone).trim();
    const cleanedSearchPhone = searchPhone.startsWith("0") ? searchPhone.substring(1) : searchPhone;

    Logger.log("📞 البحث عن رقم هاتف الطالب التجريبي: " + cleanedSearchPhone);

    const teacherIdToNameMap = new Map();
    if (teachersSheet) {
        const teachersValues = teachersSheet.getDataRange().getValues();
        teachersValues.forEach(row => {
            const teacherId = String(row[0] || '').trim();
            const teacherName = String(row[1] || '').trim();
            if (teacherId) {
                teacherIdToNameMap.set(teacherId, teacherName);
            }
        });
    }

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const trialID = String(row[0] || '').trim();
        const studentName = String(row[1] || '').trim();
        const studentAge = row[2];
        const storedPhone = String(row[3] || '').trim();
        const teacherId = String(row[4] || '').trim();
        const day = String(row[6] || '').trim();
        const time = String(row[7] || '').trim();

        const cleanedStoredPhone = storedPhone.startsWith("0") ? storedPhone.substring(1) : storedPhone;

        if (cleanedStoredPhone === cleanedSearchPhone) {
            Logger.log("Match found at row " + (i + 1));
            students.push({
                rowIndex: i + 1,
                trialID: trialID,
                name: studentName,
                age: studentAge,
                phone: storedPhone,
                teacherId: teacherId,
                teacherName: teacherIdToNameMap.get(teacherId) || teacherId,
                day: day,
                time: time,
                status: String(row[10] || '').trim() // Status (K)
            });
        }
    }

    Logger.log("✅ عدد الطلاب التجريبيين الذين تم العثور عليهم برقم الهاتف: " + students.length);
    return students;
}


// ==============================================================================
// 12. الدوال الخاصة بصفحة متابعة التجديدات والحضور والغياب
// ==============================================================================

/**
 * تجلب بيانات جميع الطلاب (المشتركين والتجريبيين) مع تفاصيل الاشتراكات والحضور/الغياب لهذا الشهر
 * لصفحة "متابعة التجديدات والحضور والغياب".
 *
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب الموحدة.
 * أو {Object} كائن خطأ.
 */
function getAllStudentsForRenewalAttendance() {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور");
  const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون"); // شيت الطلاب التجريبيين

  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود." };
  if (!subscriptionsSheet) return { error: "شيت 'الاشتراكات الحالية' غير موجود." };
  if (!teachersSheet) return { error: "شيت 'المعلمين' غير موجود." };
  if (!attendanceLogSheet) return { error: "شيت 'سجل الحضور' غير موجود." };
  if (!trialStudentsSheet) return { error: "شيت 'الطلاب التجريبيون' غير موجود." };

  const allStudentsCombined = []; // لتخزين جميع الطلاب (مشتركين وتجريبيين)

  // جلب بيانات المعلمين (Teacher ID -> Name)
  const teacherIdToNameMap = new Map();
  const teachersData = teachersSheet.getDataRange().getValues();
  teachersData.forEach(row => {
    const teacherId = String(row[0] || '').trim();
    const teacherName = String(row[1] || '').trim();
    if (teacherId) teacherIdToNameMap.set(teacherId, teacherName);
  });

  // جلب بيانات الاشتراكات (Student ID -> Subscription Details)
  const subscriptionsMap = new Map();
  const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
  subscriptionsData.forEach((row, index) => {
    if (index === 0) return; // تخطي العناوين
    const studentID = String(row[1] || '').trim(); // العمود B: Student ID
    if (studentID) {
      subscriptionsMap.set(studentID, {
        packageName: String(row[2] || '').trim(),
        teacherId: String(row[3] || '').trim(),
        subscriptionStartDate: row[4],
        subscriptionEndDate: row[5],
        attendedSessionsTotal: row[6], // عدد الحصص الحاضرة الكلي
        renewalStatus: String(row[7] || '').trim(),
        lastRenewalDate: row[8],
        totalSubscriptionAmount: row[9],
        paidAmount: row[10],
        remainingAmount: row[11]
      });
    }
  });

  // حساب الحضور والغياب للشهر الحالي
  const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();

  const studentMonthlyAttendance = new Map(); // Key: Student ID, Value: { attended: 0, absent: 0 }

  attendanceLogData.forEach((row, index) => {
    if (index === 0) return; // تخطي العناوين
    const logDate = row[4]; // العمود E: تاريخ الحصة
    const studentId = String(row[1] || '').trim(); // العمود B: Student ID/Trial ID
    const status = String(row[7] || '').trim(); // العمود H: حالة الحضور ("حضر", "غاب")

    if (logDate instanceof Date && logDate.getMonth() === currentMonth && logDate.getFullYear() === currentYear) {
      if (!studentMonthlyAttendance.has(studentId)) {
        studentMonthlyAttendance.set(studentId, { attended: 0, absent: 0 });
      }
      const counts = studentMonthlyAttendance.get(studentId);
      if (status === "حضر") {
        counts.attended++;
      } else if (status === "غاب") {
        counts.absent++;
      }
    }
  });

  // 1. معالجة الطلاب المشتركين من شيت "الطلاب"
  const studentsData = studentsSheet.getDataRange().getValues();
  studentsData.forEach((row, index) => {
    if (index === 0) return; // تخطي العناوين
    const studentID = String(row[0] || '').trim();

    const studentInfo = {
      rowIndex: index + 1,
      studentID: studentID,
      name: String(row[1] || '').trim(),
      age: row[2],
      phone: String(row[3] || '').trim(),
      studentPhone: String(row[4] || '').trim(),
      country: String(row[5] || '').trim(),
      registrationDate: row[6] ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
      basicStatus: String(row[7] || '').trim(),
      notes: String(row[8] || '').trim(),
      isTrial: false // للدلالة على أنه ليس تجريبياً
    };

    const subscriptionDetails = subscriptionsMap.get(studentID);
    if (subscriptionDetails) {
      studentInfo.packageName = subscriptionDetails.packageName;
      studentInfo.teacherId = subscriptionDetails.teacherId;
      studentInfo.teacherName = teacherIdToNameMap.get(subscriptionDetails.teacherId) || subscriptionDetails.teacherId;
      studentInfo.subscriptionStartDate = subscriptionDetails.subscriptionStartDate ? Utilities.formatDate(subscriptionDetails.subscriptionStartDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
      studentInfo.subscriptionEndDate = subscriptionDetails.subscriptionEndDate ? Utilities.formatDate(subscriptionDetails.subscriptionEndDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
      studentInfo.attendedSessionsTotal = subscriptionDetails.attendedSessionsTotal;
      studentInfo.renewalStatus = subscriptionDetails.renewalStatus;
      studentInfo.lastRenewalDate = subscriptionDetails.lastRenewalDate ? Utilities.formatDate(subscriptionDetails.lastRenewalDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
      studentInfo.totalSubscriptionAmount = subscriptionDetails.totalSubscriptionAmount;
      studentInfo.paidAmount = subscriptionDetails.paidAmount;
      studentInfo.remainingAmount = subscriptionDetails.remainingAmount;

      const totalSessionsForPackage = getTotalSessionsForPackage(studentInfo.packageName);
      studentInfo.remainingSessions = (totalSessionsForPackage > 0 && typeof studentInfo.attendedSessionsTotal === 'number') ?
                                      totalSessionsForPackage - studentInfo.attendedSessionsTotal : 'N/A';
    } else {
      // قيم افتراضية للطلاب الذين ليس لديهم اشتراك حالي (نادراً ما يحدث لمشترك)
      studentInfo.packageName = '';
      studentInfo.teacherName = '';
      studentInfo.renewalStatus = 'لم يشترك';
      studentInfo.remainingSessions = 'N/A';
      studentInfo.subscriptionEndDate = 'N/A';
      studentInfo.lastRenewalDate = 'N/A';
    }

    const monthlyCounts = studentMonthlyAttendance.get(studentID) || { attended: 0, absent: 0 };
    studentInfo.attendedThisMonth = monthlyCounts.attended;
    studentInfo.absentThisMonth = monthlyCounts.absent;

    allStudentsCombined.push(studentInfo);
  });

  // 2. معالجة الطلاب التجريبيين من شيت "الطلاب التجريبيون"
  const trialStudentsRawData = trialStudentsSheet.getDataRange().getValues();
  trialStudentsRawData.forEach((row, index) => {
    if (index === 0) return; // تخطي العناوين
    const trialID = String(row[0] || '').trim();

    const trialStudentInfo = {
      rowIndex: index + 1,
      trialID: trialID, // هنا نستخدم Trial ID
      name: String(row[1] || '').trim(),
      age: row[2],
      phone: String(row[3] || '').trim(),
      teacherId: String(row[4] || '').trim(),
      teacherName: teacherIdToNameMap.get(String(row[4] || '').trim()) || String(row[4] || '').trim(), // جلب اسم المعلم
      day1: String(row[6] || '').trim(), // اليوم الأول
      time1: String(row[7] || '').trim(), // الميعاد الأول
      registrationDate: row[8] ? Utilities.formatDate(row[8], Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
      basicStatus: String(row[10] || '').trim(), // الحالة من عمود Status في شيت التجريبيين
      isTrial: true, // للدلالة على أنه تجريبي

      // قيم افتراضية لكي تتناسب مع هيكل الجدول في الواجهة الأمامية
      studentID: '', // سيكون فارغاً للطالب التجريبي
      packageName: 'تجريبي',
      renewalStatus: 'تجريبي',
      remainingSessions: 'N/A', // أو يمكن وضع عدد الحصص التجريبية المتبقية إذا كان يتم تتبعها
      subscriptionEndDate: 'N/A',
      lastRenewalDate: 'N/A',
      attendedSessionsTotal: 0,
      totalSubscriptionAmount: 0,
      paidAmount: 0,
      remainingAmount: 0
    };

    const monthlyCounts = studentMonthlyAttendance.get(trialID) || { attended: 0, absent: 0 };
    trialStudentInfo.attendedThisMonth = monthlyCounts.attended;
    trialStudentInfo.absentThisMonth = monthlyCounts.absent;

    allStudentsCombined.push(trialStudentInfo);
  });

  Logger.log("تم جلب " + allStudentsCombined.length + " طالب لصفحة 'متابعة التجديدات والحضور والغياب'.");
  return allStudentsCombined;
}

/**
 * تسجل غياب طالب في "سجل الحضور".
 *
 * @param {string} studentId - Student ID أو Trial ID للطالب.
 * @param {string} idType - 'studentID' أو 'trialID'.
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function markAbsence(studentId, idType) {
  const attendanceLogSheet = SpreadsheetApp.getActiveSpreadheet().getSheetByName("سجل الحضور");
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!attendanceLogSheet) throw new Error("شيت 'سجل الحضور' غير موجود.");
    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!trialStudentsSheet) throw new Error("شيت 'الطلاب التجريبيون' غير موجود.");

    const today = new Date();
    const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // التحقق من عدم تسجيل الغياب مسبقًا لنفس الطالب في نفس التاريخ (يمكن تعديل هذا ليسمح بالغياب المتعدد في يوم واحد إذا كان له حصص متعددة)
    const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
    for (let i = 1; i < attendanceLogData.length; i++) {
      const logRow = attendanceLogData[i];
      const logStudentID = String(logRow[1] || '').trim(); // العمود B
      const logDateValue = logRow[4]; // العمود E: تاريخ الحصة
      const logDate = (logDateValue instanceof Date) ? Utilities.formatDate(logDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
      const logStatus = String(logRow[7] || '').trim(); // العمود H: حالة الحضور

      if (logStudentID === studentId && logDate === todayFormatted && logStatus === "غاب") {
        return { error: "تم تسجيل غياب لهذا الطالب في هذا اليوم مسبقًا." };
      }
    }

    let studentName = '';
    let teacherId = '';
    let timeSlot = '';
    let dayOfWeek = '';
    let subscriptionId = '';

    // جلب معلومات الطالب (اسم، معلم، ميعاد) بناءً على نوع الـ ID
    if (idType === 'studentID') {
        const studentData = studentsSheet.getDataRange().getValues();
        const studentRow = studentData.find(row => String(row[0] || '').trim() === studentId);
        if (studentRow) {
            studentName = String(studentRow[1] || '').trim();
            // لجلب المعلم والميعاد لطلاب الاشتراك، سنحتاج للبحث في شيت الاشتراكات والمواعيد
            const subscription = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية").getDataRange().getValues().find(subRow => String(subRow[1] || '').trim() === studentId);
            if (subscription) {
                teacherId = String(subscription[3] || '').trim(); // Teacher ID
                // لجلب اليوم والميعاد الفعليين، سنحتاج للبحث في شيت المواعيد المتاحة للمعلمين
                // هذا الجزء معقد بعض الشيء وقد لا يكون ضروريًا 100% لتسجيل الغياب فقط
                // يمكن أن نترك timeSlot و dayOfWeek فارغين أو نطلب من المعلم إدخالها عند الغياب
                timeSlot = "N/A"; // يمكن تحسين هذا
                dayOfWeek = "N/A"; // يمكن تحسين هذا
            }
        } else {
            throw new Error(`لم يتم العثور على الطالب المشترك بمعرف ${studentId}.`);
        }
    } else if (idType === 'trialID') {
        const trialStudentData = trialStudentsSheet.getDataRange().getValues();
        const trialStudentRow = trialStudentData.find(row => String(row[0] || '').trim() === studentId);
        if (trialStudentRow) {
            studentName = String(trialStudentRow[1] || '').trim();
            teacherId = String(trialStudentRow[4] || '').trim(); // Teacher ID
            dayOfWeek = String(trialStudentRow[6] || '').trim(); // اليوم
            timeSlot = String(trialStudentRow[7] || '').trim(); // الميعاد
        } else {
            throw new Error(`لم يتم العثور على الطالب التجريبي بمعرف ${studentId}.`);
        }
    } else {
        throw new Error("نوع معرف الطالب غير صالح.");
    }

    const attendanceId = generateUniqueAttendanceId(attendanceLogSheet); // دالة موجودة

    // الأعمدة: Attendance ID, Student ID, Teacher ID, Subscription ID, تاريخ الحصة, وقت الحصة, اليوم, حالة الحضور, نوع الحصة, ملاحظات المعلم
    attendanceLogSheet.appendRow([
      attendanceId,   // A
      studentId,      // B
      teacherId,      // C
      subscriptionId, // D (فارغ إذا كان تجريبياً أو غير متاح للمشتركين هنا)
      today,          // E
      timeSlot,       // F
      dayOfWeek,      // G
      "غاب",          // H (حالة الحضور)
      (idType === 'trialID' ? "تجريبية" : "عادية"), // I (نوع الحصة)
      ""              // J (ملاحظات المعلم)
    ]);
    Logger.log(`تم تسجيل غياب لـ ${studentName} (ID: ${studentId}) في ${todayFormatted}.`);

    return { success: `تم تسجيل غياب للطالب ${studentName} بنجاح.` };

  } catch (e) {
    Logger.log("خطأ في markAbsence: " + e.message);
    return { error: `فشل تسجيل الغياب: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}



// ==============================================================================
// 13. الدوال الخاصة بالتكامل مع شيتات المعلمين (تُستدعى من شيت المعلم)
// ==============================================================================

/**
 * تستقبل بيانات حضور/غياب طالب من شيت المعلم وتُحدث سجل الحضور المركزي
 * وحالة الاشتراكات في ملف المشرف.
 *
 * @param {string} teacherId - Teacher ID للمعلم الذي سجل الحضور/الغياب.
 * @param {string} studentId - Student ID أو Trial ID للطالب.
 * @param {string} studentName - اسم الطالب.
 * @param {string} day - يوم الحصة.
 * @param {string} timeSlot - ميعاد الحصة.
 * @param {string} status - حالة الحضور ("حضر" أو "غاب").
 * @returns {Object} رسالة نجاح أو خطأ.
 */
function updateSupervisorAttendanceSheet(teacherId, studentId, studentName, day, timeSlot, status) {
  const attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب"); // لجلب اسم الطالب المشترك
  const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون"); // لجلب اسم الطالب التجريبي

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!attendanceLogSheet) throw new Error("شيت 'سجل الحضور' غير موجود في ملف المشرف.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود في ملف المشرف.");
    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود في ملف المشرف.");
    if (!trialStudentsSheet) throw new Error("شيت 'الطلاب التجريبيون' غير موجود في ملف المشرف.");

    const today = new Date();
    const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // التحقق من عدم تسجيل الحضور/الغياب مسبقًا لنفس الطالب في نفس التاريخ والميعاد
    // هذا التحقق يمنع التكرار إذا تم الإرسال أكثر من مرة لنفس الحصة.
    const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
    for (let i = 1; i < attendanceLogData.length; i++) {
        const logRow = attendanceLogData[i];
        const logStudentID = String(logRow[1] || '').trim(); // العمود B
        const logTeacherID = String(logRow[2] || '').trim(); // العمود C
        const logDateValue = logRow[4]; // العمود E: تاريخ الحصة
        const logDate = (logDateValue instanceof Date) ? Utilities.formatDate(logDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        const logTimeSlot = String(logRow[5] || '').trim(); // العمود F
        const logStatus = String(logRow[7] || '').trim(); // العمود H

        if (logStudentID === studentId && logTeacherID === teacherId && logDate === todayFormatted && logTimeSlot === timeSlot && logStatus === status) {
            return { error: `تم تسجيل ${status} لهذا الطالب في هذا الميعاد وهذا اليوم مسبقًا في سجل المشرف.` };
        }
    }


    // تحديد نوع الحصة (عادية / تجريبية)
    let classType = (studentId.startsWith("TRL")) ? "تجريبية" : "عادية";
    let subscriptionId = '';
    let packageName = '';
    let renewalStatus = '';
    let totalPackageSessions = 0;
    let subscriptionRowIndex = -1;

    // جلب بيانات الاشتراك فقط للطلاب المشتركين لتحديثهم
    if (classType === "عادية") {
        const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
        for (let i = 1; i < subscriptionsData.length; i++) {
            if (String(subscriptionsData[i][1] || '').trim() === studentId) { // العمود B: Student ID
                subscriptionRowIndex = i;
                subscriptionId = String(subscriptionsData[i][0] || '').trim(); // Subscription ID
                packageName = String(subscriptionsData[i][2] || '').trim(); // اسم الباقة
                renewalStatus = String(subscriptionsData[i][7] || '').trim(); // الحالة التفصيلية للتجديد
                totalPackageSessions = getTotalSessionsForPackage(packageName); // دالة موجودة
                break;
            }
        }
        if (subscriptionRowIndex === -1) {
            // إذا لم يتم العثور على اشتراك لطالب مشترك، قد يكون هذا خطأ أو حالة استثنائية
            Logger.log(`تحذير: لم يتم العثور على اشتراك للطالب المشترك ID ${studentId} لتحديثه في سجل المشرف.`);
            // سنستمر في تسجيل الحضور/الغياب لكن لن نحدث الاشتراك
        }
    }

    // 1. تسجيل الحضور/الغياب في شيت "سجل الحضور" المركزي
    const attendanceId = generateUniqueAttendanceId(attendanceLogSheet); // دالة موجودة
    attendanceLogSheet.appendRow([
      attendanceId,   // A: Attendance ID
      studentId,      // B: Student ID / Trial ID
      teacherId,      // C: Teacher ID
      subscriptionId, // D: Subscription ID (فارغ إذا كان تجريبياً)
      today,          // E: تاريخ الحصة
      timeSlot,       // F: وقت الحصة
      day,            // G: اليوم
      status,         // H: حالة الحضور ("حضر" أو "غاب")
      classType,      // I: نوع الحصة ("عادية" أو "تجريبية")
      ""              // J: ملاحظات المعلم
    ]);
    Logger.log(`تم تسجيل ${status} للطالب ${studentName} (ID: ${studentId}) في سجل المشرف.`);

    // 2. تحديث عدد الحصص الحاضرة في شيت "الاشتراكات الحالية" (فقط للحصص العادية التي تم "حضر" فيها)
    if (classType === "عادية" && status === "حضر" && subscriptionRowIndex !== -1) {
      const currentAttendedSessionsCell = subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7); // العمود G: عدد الحصص الحاضرة
      let currentSessions = currentAttendedSessionsCell.getValue();
      currentSessions = (typeof currentSessions === 'number') ? currentSessions : 0;
      subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7).setValue(currentSessions + 1);
      Logger.log(`تم تحديث عدد الحصص الحاضرة للطالب ${studentId} إلى ${currentSessions + 1} في سجل المشرف.`);

      // تحديث "الحالة التفصيلية للتجديد" إذا وصل الطالب إلى الحد الأقصى من الحصص
      if (totalPackageSessions > 0 && (currentSessions + 1) >= totalPackageSessions) {
        subscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue("يحتاج للتجديد"); // العمود H
        Logger.log(`حالة تجديد الطالب ${studentId} تم تحديثها إلى "يحتاج للتجديد" في سجل المشرف.`);
      }
    }

    return { success: `تم تحديث سجل المشرف بحالة ${status} للطالب ${studentName}.` };

  } catch (e) {
    Logger.log("خطأ في updateSupervisorAttendanceSheet: " + e.message);
    return { error: `فشل تحديث سجل المشرف: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


// ==============================================================================
// 15. الدوال الخاصة بصفحة الطلاب الذين يحتاجون للتجديد (في ملف المشرف)
// ==============================================================================

/**
 * تجلب قائمة الطلاب الذين يحتاجون للتجديد بناءً على معايير:
 * 1. استنفذ عدد الحصص المقررة في اشتراكه (remainingSessions <= 0).
 * 2. مر شهر وخمسة أيام على آخر تاريخ تجديد (lastRenewalDate + 35 days).
 *
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب الذين يحتاجون للتجديد.
 * أو {Object} كائن خطأ.
 */
function getStudentsWhoNeedRenewal() {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");

  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود." };
  if (!subscriptionsSheet) return { error: "شيت 'الاشتراكات الحالية' غير موجود." };
  if (!teachersSheet) return { error: "شيت 'المعلمين' غير موجود." };
  if (!packagesSheet) return { error: "شيت 'الباقات' غير موجود." };

  const studentsNeedingRenewal = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0); // لضمان المقارنة باليوم فقط

  // جلب بيانات المعلمين (ID -> Name)
  const teacherIdToNameMap = new Map();
  teachersSheet.getDataRange().getValues().forEach(row => {
    const teacherId = String(row[0] || '').trim();
    const teacherName = String(row[1] || '').trim();
    if (teacherId) teacherIdToNameMap.set(teacherId, teacherName);
  });

  // جلب بيانات الطلاب الأساسية (ID -> Student Info)
  const studentsMap = new Map();
  studentsSheet.getDataRange().getValues().forEach((row, index) => {
    if (index === 0) return;
    const studentID = String(row[0] || '').trim();
    if (studentID) {
      studentsMap.set(studentID, {
        name: String(row[1] || '').trim(),
        phone: String(row[3] || '').trim(),
        basicStatus: String(row[7] || '').trim() // الحالة الأساسية
      });
    }
  });

  // جلب بيانات الاشتراكات والباقات
  const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
  const packagesData = packagesSheet.getDataRange().getValues();
  const packageTotalSessionsMap = new Map(); // Key: packageName, Value: totalSessions
  packagesData.forEach((row, index) => {
    if (index === 0) return;
    const pkgName = String(row[0] || '').trim();
    const totalSessions = row[2]; // العمود C
    if (pkgName && typeof totalSessions === 'number') {
      packageTotalSessionsMap.set(pkgName, totalSessions);
    }
  });


  subscriptionsData.forEach((row, index) => {
    if (index === 0) return; // تخطي العناوين
    
    const studentID = String(row[1] || '').trim(); // العمود B
    const packageName = String(row[2] || '').trim(); // العمود C
    const teacherId = String(row[3] || '').trim(); // العمود D
    const lastRenewalDate = row[8]; // العمود I: تاريخ آخر تجديد
    const attendedSessions = row[6]; // العمود G: عدد الحصص الحاضرة
    const renewalStatus = String(row[7] || '').trim(); // العمود H: الحالة التفصيلية للتجديد

    const studentInfo = studentsMap.get(studentID);
    if (!studentInfo) return; // تخطي إذا لم يتم العثور على معلومات الطالب الأساسية

    // تخطي الطلاب التجريبيين أو الذين لا يجب تجديدهم من هنا
    if (renewalStatus === "تجريبي" || renewalStatus === "لم يشترك" || studentInfo.basicStatus === "تجريبي") {
      return;
    }

    let needsRenewalBySessions = false;
    let needsRenewalByDate = false;
    let reasonForRenewal = '';

    const totalSessions = packageTotalSessionsMap.get(packageName) || 0;
    const remainingSessions = (totalSessions > 0 && typeof attendedSessions === 'number') ? totalSessions - attendedSessions : null;

    // الشرط 1: استنفاد الحصص
    if (remainingSessions !== null && remainingSessions <= 0 && totalSessions > 0) {
      needsRenewalBySessions = true;
      reasonForRenewal = 'بسبب استنفاد الحصص';
    }
    // الشرط 2: مرور شهر و خمسة أيام على آخر تجديد
    if (lastRenewalDate instanceof Date) {
      const renewalDueDate = new Date(lastRenewalDate);
      renewalDueDate.setDate(renewalDueDate.getDate() + 35); // إضافة 35 يومًا (شهر + 5 أيام)
      renewalDueDate.setHours(0, 0, 0, 0); // لضمان المقارنة باليوم فقط

      if (today >= renewalDueDate) {
        needsRenewalByDate = true;
        if (reasonForRenewal) { // إذا كان هناك سبب آخر
          reasonForRenewal += '، وبسبب انتهاء المدة';
        } else {
          reasonForRenewal = 'بسبب انتهاء المدة';
        }
      }
    }

    // إذا كان يحتاج للتجديد بأي من السببين، نضيفه إلى القائمة
    if (needsRenewalBySessions || needsRenewalByDate) {
      studentsNeedingRenewal.push({
        name: studentInfo.name,
        studentID: studentID,
        phone: studentInfo.phone,
        teacherName: teacherIdToNameMap.get(teacherId) || teacherId,
        packageName: packageName,
        renewalStatus: renewalStatus, // الحالة الحالية (مثلاً: يحتاج للتجديد، انتهت المدة، تجاوز الحد)
        reasonForRenewal: reasonForRenewal,
        remainingSessions: remainingSessions !== null ? remainingSessions : 'غير محدد',
        lastRenewalDate: lastRenewalDate ? Utilities.formatDate(lastRenewalDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : 'N/A'
      });
    }
  });

  Logger.log("تم جلب " + studentsNeedingRenewal.length + " طالب يحتاجون للتجديد.");
  return studentsNeedingRenewal;
}




// ==============================================================================
// 16. الدوال الخاصة بتسجيل الدفعات المالية (في ملف المشرف)
// ==============================================================================

/**
 * دالة لتوليد معرف دفعة جديد وفريد بناءً على آخر معرف في شيت "سجل الدفعات".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} paymentRecordsSheet - شيت "سجل الدفعات".
 * @returns {string} معرف الدفعة الجديد (مثال: PAY001).
 */
function generateUniquePaymentId(paymentRecordsSheet) {
  const lastRow = paymentRecordsSheet.getLastRow();
  let lastGeneratedIdNum = 0;
  if (lastRow >= 2) { // نبدأ من الصف الثاني لتخطي العناوين
    const paymentIds = paymentRecordsSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // العمود A لـ Payment ID
    const numericIds = paymentIds.flat().map(id => {
      const numPart = String(id).replace('PAY', '');
      return parseInt(numPart) || 0;
    }).filter(Number);
    lastGeneratedIdNum = numericIds.length > 0 ? Math.max(...numericIds) : 0;
  }
  return `PAY${(lastGeneratedIdNum + 1).toString().padStart(3, '0')}`;
}

/**
 * تسجل دفعة مالية جديدة في شيت "سجل الدفعات".
 *
 * @param {Object} paymentData - كائن يحتوي على:
 * - studentID: معرف الطالب.
 * - subscriptionId: معرف الاشتراك المرتبط (اختياري، يمكن أن يكون فارغًا).
 * - paymentDate: تاريخ الدفعة (بتنسيق yyyy-MM-dd).
 * - paymentAmount: المبلغ المدفوع.
 * - paymentMethod: طريقة الدفع.
 * - paymentNotes: ملاحظات إضافية.
 * @returns {Object} رسالة نجاح أو خطأ.
 */
function recordPayment(paymentData) {
  const paymentRecordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الدفعات");

  if (!paymentRecordsSheet) return { error: "شيت 'سجل الدفعات' غير موجود." };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const newPaymentId = generateUniquePaymentId(paymentRecordsSheet);

    // الأعمدة الجديدة: Payment ID - Subscription ID - Student ID - تاريخ الدفع - المبلغ المدفوع - طريقة الدفع - ملاحظات
    const newRow = [
      newPaymentId,
      paymentData.subscriptionId || '', // Subscription ID
      paymentData.studentID,
      paymentData.paymentDate,
      paymentData.paymentAmount,
      paymentData.paymentMethod,
      paymentData.paymentNotes
    ];

    paymentRecordsSheet.appendRow(newRow);
    Logger.log(`تم تسجيل دفعة ${newPaymentId} للطالب ${paymentData.studentID} بمبلغ ${paymentData.paymentAmount}.`);
    return { success: "تم تسجيل الدفعة بنجاح." };

  } catch (e) {
    Logger.log("خطأ في recordPayment: " + e.message);
    return { error: "فشل تسجيل الدفعة: " + e.message };
  } finally {
    lock.releaseLock();
  }
}



/**
 * يجلب سجل الدفعات لطالب محدد من شيت "سجل الدفعات".
 *
 * @param {string} studentID - معرف الطالب.
 * @returns {Array<Object>} مصفوفة من كائنات سجل الدفعات.
 * أو {Object} كائن خطأ.
 */
function getPaymentRecords(studentID) {
  const paymentRecordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الدفعات");

  if (!paymentRecordsSheet) return { error: "شيت 'سجل الدفعات' غير موجود." };

  try {
    const paymentRecordsData = paymentRecordsSheet.getDataRange().getValues();
    const studentPayments = [];

    for (let i = 1; i < paymentRecordsData.length; i++) {
      // الأعمدة الجديدة: Payment ID (0), Subscription ID (1), Student ID (2), تاريخ الدفع (3), المبلغ المدفوع (4), طريقة الدفع (5), ملاحظات (6)
      if (String(paymentRecordsData[i][2] || '').trim() === String(studentID).trim()) { // العمود C: Student ID
        studentPayments.push({
          paymentId: paymentRecordsData[i][0],
          subscriptionId: paymentRecordsData[i][1],
          studentId: paymentRecordsData[i][2],
          paymentDate: paymentRecordsData[i][3] ? Utilities.formatDate(paymentRecordsData[i][3], Session.getScriptTimeZone(), "yyyy-MM-dd") : '',
          paymentAmount: paymentRecordsData[i][4],
          paymentMethod: paymentRecordsData[i][5],
          paymentNotes: paymentRecordsData[i][6]
        });
      }
    }
    return studentPayments;

  } catch (e) {
    Logger.log("خطأ في getPaymentRecords: " + e.message);
    return { error: "فشل جلب سجل الدفعات: " + e.message };
  }
}


// ==============================================================================
// 14. الدوال الخاصة بتجديد اشتراكات الطلاب (في ملف المشرف)
// ==============================================================================

/**
 * تُعالج طلب تجديد اشتراك طالب، وتُحدث بيانات الاشتراك في شيت "الاشتراكات الحالية".
 *
 * @param {Object} renewalData - كائن يحتوي على:
 * - studentID: معرف الطالب.
 * - subscriptionId: معرف الاشتراك الحالي.
 * - packageName: اسم الباقة الجديدة.
 * - paymentStatus: حالة الدفع للتجديد (تم الدفع، لم يتم الدفع، تم دفع جزء).
 * - renewalNotes: ملاحظات إضافية.
 * @returns {Object} رسالة نجاح أو خطأ.
 */
function processStudentSubscriptionRenewal(renewalData) {
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب"); // لتحديث الحالة الأساسية للطالب إذا تغيرت

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");
    if (!packagesSheet) throw new Error("شيت 'الباقات' غير موجود.");
    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");

    const today = new Date();
    const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
    let subscriptionRowIndex = -1;

    // البحث عن صف الاشتراك للطالب
    for (let i = 1; i < subscriptionsData.length; i++) {
      if (String(subscriptionsData[i][1] || '').trim() === String(renewalData.studentID).trim()) { // العمود B: Student ID
        subscriptionRowIndex = i;
        break;
      }
    }

    if (subscriptionRowIndex === -1) {
      throw new Error(`لم يتم العثور على اشتراك حالي للطالب ID ${renewalData.studentID}.`);
    }

    // جلب تفاصيل الباقة الجديدة
    const packageDetails = getPackageDetails(renewalData.packageName); // دالة موجودة
    if (!packageDetails) {
      throw new Error(`لم يتم العثور على تفاصيل الباقة: ${renewalData.packageName}.`);
    }

    // تحديث بيانات الاشتراك
    // الأعمدة في شيت "الاشتراكات الحالية":
    // A: Subscription ID, B: Student ID, C: اسم الباقة, D: Teacher ID, E: تاريخ بداية الاشتراك, F: تاريخ نهاية الاشتراك المتوقع,
    // G: عدد الحصص الحاضرة, H: الحالة التفصيلية للتجديد, I: تاريخ آخر تجديد, J: مبلغ الاشتراك الكلي,
    // K: المبلغ المدفوع حتى الآن, L: المبلغ المتبقي, M: Absences Count (غيابات غير مخصومة), N: ملاحظات خاصة بالاشتراك

    // حساب تاريخ نهاية الاشتراك المتوقع (بالنسبة للاشتراكات الشهرية/نصف سنوية)
    let endDate = "";
    let subscriptionType = packageDetails['نوع الباقة'];

    const newStartDate = new Date(today); // تاريخ بداية الاشتراك الجديد هو اليوم
    if (subscriptionType === "شهري") {
        newStartDate.setMonth(newStartDate.getMonth() + 1);
        endDate = Utilities.formatDate(newStartDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (subscriptionType === "نصف سنوي") {
        newStartDate.setMonth(newStartDate.getMonth() + 6);
        endDate = Utilities.formatDate(newStartDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else { // لأي نوع باقة أخرى، يمكن تحديد منطق مختلف
        endDate = 'غير محدد';
    }


    // تحديد حالة التجديد والمبالغ
    let newRenewalStatus = "تم التجديد";
    let paidAmount = 0;
    let remainingAmount = packageDetails['السعر'];
    const ABSENCES_COUNT_COL_INDEX = 13; // العمود M

    if (renewalData.paymentStatus === "تم الدفع") {
        paidAmount = packageDetails['السعر'];
        remainingAmount = 0;
    } else if (renewalData.paymentStatus === "لم يتم الدفع") {
        paidAmount = 0;
        remainingAmount = packageDetails['السعر'];
        newRenewalStatus = "لم يتم الدفع";
    } else if (renewalData.paymentStatus === "تم دفع جزء") {
        paidAmount = 0; // يمكن تعديلها إذا تم تمرير المبلغ المدفوع جزئياً
        remainingAmount = packageDetails['السعر'];
        newRenewalStatus = "تم دفع جزء";
    }

    // تحديث الصف في شيت الاشتراكات الحالية
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 3).setValue(renewalData.packageName); // C: اسم الباقة
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 5).setValue(todayFormatted); // E: تاريخ بداية الاشتراك (تاريخ التجديد)
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 6).setValue(endDate); // F: تاريخ نهاية الاشتراك المتوقع
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7).setValue(0); // G: عدد الحصص الحاضرة (يتم تصفيرها)
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue(newRenewalStatus); // H: الحالة التفصيلية للتجديد
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 9).setValue(todayFormatted); // I: تاريخ آخر تجديد (هو اليوم)
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 10).setValue(packageDetails['السعر']); // J: مبلغ الاشتراك الكلي
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 11).setValue(paidAmount); // K: المبلغ المدفوع حتى الآن
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 12).setValue(remainingAmount); // L: المبلغ المتبقي
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, ABSENCES_COUNT_COL_INDEX).setValue(0); // M: تصفير عداد الغيابات غير المخصومة
    subscriptionsSheet.getRange(subscriptionRowIndex + 1, 14).setValue(renewalData.renewalNotes); // N: ملاحظات خاصة بالاشتراك (إذا كان العمود N هو الملاحظات)

    Logger.log(`تم تجديد اشتراك الطالب ${renewalData.studentID} بنجاح.`);

    // تحديث الحالة الأساسية للطالب في شيت "الطلاب" إلى "مشترك"
    if (renewalData.paymentStatus === "تم الدفع") {
        const studentsData = studentsSheet.getDataRange().getValues();
        const studentRow = studentsData.find(row => String(row[0] || '').trim() === renewalData.studentID);
        if (studentRow) {
            const studentSheetRowIndex = studentsData.indexOf(studentRow) + 1;
            studentsSheet.getRange(studentSheetRowIndex, 8).setValue("مشترك"); // العمود H: الحالة الأساسية للطالب
            Logger.log(`تم تحديث الحالة الأساسية للطالب ${renewalData.studentID} إلى "مشترك".`);
        }
    }


    // تسجيل الدفعة المرتبطة بالتجديد (إذا تم الدفع)
    if (renewalData.paymentStatus !== "لم يتم الدفع") {
        const paymentData = {
            studentID: renewalData.studentID,
            subscriptionId: subscriptionsSheet.getRange(subscriptionRowIndex + 1, 1).getValue(), // <--- جديد: جلب Subscription ID من العمود A
            paymentDate: todayFormatted,
            paymentAmount: paidAmount, // المبلغ المدفوع
            paymentMethod: renewalData.paymentStatus, // طريقة الدفع (يمكن تعديلها)
            paymentNotes: `تجديد الاشتراك. ${renewalData.renewalNotes}` // ملاحظات التجديد
        };
        const paymentResult = recordPayment(paymentData);
        if (paymentResult.error) {
            Logger.log("تحذير: فشل تسجيل الدفعة المرتبطة بالتجديد: " + paymentResult.error);
        } else {
            Logger.log(`تم تسجيل دفعة مرتبطة بتجديد الطالب ${renewalData.studentID}.`);
        }
    }

    return { success: `تم تجديد اشتراك الطالب ${renewalData.studentID} بنجاح.` };


  } catch (e) {
    Logger.log("خطأ في processStudentSubscriptionRenewal: " + e.message);
    return { error: `فشل تجديد الاشتراك: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}




/**
 * دالة مساعدة لتحويل تنسيق وقت نص عادي (من الشيتات القديمة) إلى تنسيق رأس الميعاد القياسي (HH:mm - HH:mm).
 *
 * @param {string} oldPlainTimeFormat - سلسلة الوقت بالتنسيق القديم (مثلاً "9:00 ص", "1:30 م", "8:30", "p 3").
 * @returns {string} الميعاد بتنسيق رأس العمود (مثلاً "09:00 - 09:30").
 */
function convertOldPlainTimeFormatToHeaderFormat(oldPlainTimeFormat) {
  if (typeof oldPlainTimeFormat !== 'string' || oldPlainTimeFormat.trim() === '') {
    return ''; // يرجع سلسلة فارغة لو الوقت غير صالح
  }

  // 1. تحويل الوقت المدخل إلى تنسيق 24 ساعة (HH:mm)
  const startTime24hr = convertTo24HourFormat(oldPlainTimeFormat);

  // 2. حساب وقت النهاية (+30 دقيقة)
  const [hours, minutes] = startTime24hr.split(':').map(Number);
  if (isNaN(hours) || isNaN(minutes)) {
    Logger.log("Warning: Invalid time part after conversion: " + startTime24hr);
    return '';
  }

  const startDate = new Date(); // نستخدم تاريخ اليوم مؤقتاً للحسابات
  startDate.setHours(hours, minutes, 0, 0);

  const endDate = new Date(startDate.getTime() + 30 * 60 * 1000); // إضافة 30 دقيقة

  const endHours = endDate.getHours();
  const endMinutes = endDate.getMinutes();

  const endTime24hr = `${String(endHours).padStart(2, '0')}:${String(endMinutes).padStart(2, '0')}`;

  // 3. بناء سلسلة رأس العمود
  return `${startTime24hr} - ${endTime24hr}`;
}


// ==============================================================================
// 10. الدوال الخاصة بنقل البيانات (Migration Functions)
// ==============================================================================

/**
 * دالة رئيسية لنقل بيانات الطلاب من شيت "بيانات الطلبة" القديم (بناءً على الثوابت الجديدة)
 * إلى هيكل الشيتات الجديد (الطلاب، الاشتراكات الحالية، المواعيد المتاحة للمعلمين).
 *
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function migrateStudentsOnly() {
  // === قم بتغيير هذا الـ ID بالـ ID الخاص بملف Google Sheets القديم الجديد هذا ===
  // تأكد أن هذا هو ID ملف Google Sheet القديم الذي يحتوي على MASTER_SHEET_NAME و TEACHERS_SHEET_NAME
  const OLD_SPREADSHEET_ID_NEW_STRUCTURE = "1XWRFKp-LM7detp42X4bqjhVgkaJT00of6FWvtx8sYL0";
  // ====================================================================

  let oldDataSpreadsheets;
  try {
    oldDataSpreadsheets = SpreadsheetApp.openById(OLD_SPREADSHEET_ID_NEW_STRUCTURE);
  } catch (e) {
    return { error: `فشل فتح ملف Google Sheets القديم بالـ ID المقدم: ${OLD_SPREADSHEET_ID_NEW_STRUCTURE}. الخطأ: ${e.message}` };
  }

  const oldMasterSheet = oldDataSpreadsheets.getSheetByName(MASTER_SHEET_NAME);
  const oldTeachersSheet = oldDataSpreadsheets.getSheetByName(TEACHERS_SHEET_NAME);
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const newTeachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(300000); // زيادة وقت القفل إلى 5 دقائق لعملية النقل

    // التحقق من وجود الشيتات الضرورية
    if (!oldMasterSheet) throw new Error(`شيت '${MASTER_SHEET_NAME}' القديم غير موجود في الملف القديم.`);
    if (!oldTeachersSheet) throw new Error(`شيت '${TEACHERS_SHEET_NAME}' القديم غير موجود في الملف القديم.`);
    if (!studentsSheet) throw new Error("شيت 'الطلاب' الجديد غير موجود.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' الجديد غير موجود.");
    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' الجديد غير موجود.");
    if (!newTeachersSheet) throw new Error("شيت 'المعلمين' الجديد غير موجود.");
    if (!packagesSheet) throw new Error("شيت 'الباقات' الجديد غير موجود.");


    const oldMasterData = oldMasterSheet.getDataRange().getValues();
    if (oldMasterData.length < MASTER_FIRST_DATA_ROW) {
      return { success: "لا توجد بيانات طلاب في الشيت الرئيسي القديم لنقلها." };
    }

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let migratedCount = 0;
    let failedCount = 0;
    const errors = [];

    // 1. بناء خريطة (اسم المعلم -> Teacher ID) من شيت "المعلمين" في الملف القديم
    const oldTeachersData = oldTeachersSheet.getDataRange().getValues();
    const oldTeacherNameToIdMap = new Map();
    const OLD_TEACHERS_SHEET_TEACHER_NAME_COL_INDEX = TEACHER_NAME_COL - 1; // العمود A هو 0-based
    const OLD_TEACHERS_SHEET_TEACHER_ID_COL_INDEX = TEACHER_ID_COL - 1; // العمود A هو 0-based - <--- هذا هو المؤشر الصحيح لـ Teacher ID في OLD Teachers sheet
    // Logger.log("Using OLD_TEACHERS_SHEET_TEACHER_NAME_COL_INDEX: " + OLD_TEACHERS_SHEET_TEACHER_NAME_COL_INDEX); // Debug
    // Logger.log("Using OLD_TEACHERS_SHEET_TEACHER_ID_COL_INDEX: " + OLD_TEACHERS_SHEET_TEACHER_ID_COL_INDEX); // Debug

    for (let i = TEACHER_FIRST_DATA_ROW - 1; i < oldTeachersData.length; i++) {
        const teacherName = String(oldTeachersData[i][OLD_TEACHERS_SHEET_TEACHER_NAME_COL_INDEX] || '').trim();
        const teacherId = String(oldTeachersData[i][OLD_TEACHERS_SHEET_TEACHER_ID_COL_INDEX] || '').trim();
        if (teacherName && teacherId) {
            oldTeacherNameToIdMap.set(teacherName, teacherId);
        }
    }
    Logger.log("Old Teacher Name to ID Map: " + JSON.stringify(Array.from(oldTeacherNameToIdMap.entries())));


    // 2. جلب بيانات الباقات من الشيت الجديد (لا يوجد تحويل للأسماء)
    const packagesMap = new Map();
    const packagesData = packagesSheet.getDataRange().getValues();
    for (let i = 1; i < packagesData.length; i++) {
        const packageName = String(packagesData[i][0] || '').trim();
        if (packageName) {
            packagesMap.set(packageName, {
                'اسم الباقة': packageName,
                'مدة الحصة (دقيقة)': packagesData[i][1],
                'عدد الحصص الكلي': packagesData[i][2],
                'السعر': packagesData[i][3],
                'نوع الباقة': String(packagesData[i][4] || '').trim()
            });
        }
    }

    // 3. قراءة جميع بيانات المواعيد المتاحة للمعلمين من الشيت الجديد مرة واحدة (للكتابة بالجملة)
    const allTeacherSlotsDataValues = teachersAvailableSlotsSheet.getDataRange().getValues();
    const allTeacherSlotsHeaders = allTeacherSlotsDataValues[0];
    const timeSlotHeaderToColIndexMap = new Map();
    const startColIndexForSlots = 2;
    for (let i = startColIndexForSlots; i < allTeacherSlotsHeaders.length; i++) {
        const header = String(allTeacherSlotsHeaders[i] || '').trim();
        if (header) {
            timeSlotHeaderToColIndexMap.set(header, i);
        }
    }
    const mutableTeachersAvailableSlotsData = allTeacherSlotsDataValues.map(row => [...row]);


    const newStudentsRows = [];
    const newSubscriptionsRows = [];


    // 4. المرور على كل صف طالب في الشيت الرئيسي القديم (بيانات الطلبة)
    for (let i = MASTER_FIRST_DATA_ROW - 1; i < oldMasterData.length; i++) {
      const oldRow = oldMasterData[i];
      try {
        const oldStudentID = String(oldRow[MASTER_STUDENT_ID_COL - 1] || '').trim();
        const oldStudentName = String(oldRow[MASTER_STUDENT_NAME_COL - 1] || '').trim();
        const oldAge = oldRow[MASTER_STUDENT_AGE_COL - 1];
        const oldPhoneNumber = String(oldRow[MASTER_STUDENT_NUMBER_COL - 1] || '').trim();

        // **التعديل هنا: قراءة اسم المعلم من العمود E في MASTER_SHEET_NAME (المؤشر 4)**
        const oldTeacherName = String(oldRow[4] || '').trim(); // <--- هذا هو التعديل الأساسي لقراءة اسم المعلم من العمود E في MASTER_SHEET_NAME

        const oldSubscriptionType = String(oldRow[MASTER_SUB_TYPE_COL - 1] || '').trim();
        const oldSystem = String(oldRow[MASTER_SYSTEM_COL - 1] || '').trim();
        const oldTotalAttendance = oldRow[MASTER_TOTAL_ATTENDANCE_COL - 1];
        const oldRenewalStatus = String(oldRow[MASTER_RENEWAL_STATUS_COL - 1] || '').trim();
        const oldLastPaymentDate = oldRow[MASTER_LAST_PAYMENT_DATE_COL - 1];
        const oldAmount = oldRow[MASTER_AMOUNT_COL - 1];
        const oldSubscriptionDate = oldRow[MASTER_SUBSCRIPTION_DATE_COL - 1];

        // المواعيد من الشيت القديم (بيانات الطلبة)
        // بناءً على تنسيق الشيت القديم الذي كان فيه اسم المعلم في D والمواعيد في E و F و G و H
        // بما أن اسم المعلم الآن في E، فالمواعيد ستكون في أعمدة مختلفة.
        // يجب أن تؤكد أعمدة المواعيد في MASTER_SHEET_NAME.
        // بناءً على سياق المشاكل السابقة، يبدو أنها كانت تُقرأ من المؤشرين 5 و 7 (أعمدة F و H) في الصف.
        // سنفترض أنها الأعمدة F,G,H,I إذا كان E للمعلم.
        // إذا كان العمود E هو اسم المعلم، فإن المواعيد (اليوم1، الميعاد1، اليوم2، الميعاد2) ستكون في F, G, H, I.
        // أي المؤشرات 5, 6, 7, 8 في الصف (0-based)
        const oldDay1 = String(oldRow[5] || '').trim(); // العمود F (مؤشر 5)
        const oldTime1 = oldRow[6]; // العمود G (مؤشر 6) - قد تكون كائن Date
        const oldDay2 = String(oldRow[7] || '').trim(); // العمود H (مؤشر 7)
        const oldTime2 = oldRow[8]; // العمود I (مؤشر 8) - قد تكون كائن Date


        if (!oldStudentName) {
            Logger.log(`Skipping empty student name in row ${i + 1} of old sheet.`);
            continue;
        }
        if (!oldStudentID) {
            errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}): Missing Student ID in old sheet. Skipping student.`);
            failedCount++;
            continue;
        }

        // 5. جلب Teacher ID من الخريطة المستخرجة من شيت المعلمين القديم
        const teacherId = oldTeacherNameToIdMap.get(String(oldTeacherName).trim()); // <--- استخدام oldTeacherName مباشرة
        
        if (!teacherId) {
            errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}): Teacher '${oldTeacherName}' not found in OLD Teachers sheet (column A) to get ID. Skipping student.`); // <--- رسالة خطأ محدثة
            failedCount++;
            continue;
        }


        // 6. تحديد حالة الطالب الأساسية في الشيت الجديد
        let studentBasicStatus = "معلق";
        if (oldRenewalStatus === "تم التجديد") {
            studentBasicStatus = "مشترك";
        } else if (oldRenewalStatus === "تجريبي") {
            studentBasicStatus = "تجريبي";
        }


        let endDate = "";
        let subscriptionType = oldSubscriptionType;

        if (subscriptionType.includes("شهري")) {
            const startDate = new Date(today);
            startDate.setMonth(startDate.getMonth() + 1);
            endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (subscriptionType.includes("نصف سنوي")) {
            const startDate = new Date(today);
            startDate.setMonth(startDate.getMonth() + 6);
            endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (subscriptionType.includes("سنوي")) {
            const startDate = new Date(today);
            startDate.setFullYear(startDate.getFullYear() + 1);
            endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
            endDate = 'غير محدد';
        }

        let paidAmount = 0;
        let remainingAmount = oldAmount || 0;

        if (oldRenewalStatus === "تم الدفع") {
            paidAmount = oldAmount || 0;
            remainingAmount = 0;
        } else if (oldRenewalStatus === "تم دفع جزء") {
            paidAmount = 0;
        }

        // 9. إضافة الصفوف إلى قوائم الكتابة بالجملة
        newStudentsRows.push([
          oldStudentID,
          oldStudentName,
          oldAge,
          oldPhoneNumber,
          "",
          "",
          Utilities.formatDate(oldSubscriptionDate instanceof Date ? oldSubscriptionDate : new Date(oldSubscriptionDate), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          studentBasicStatus,
          ""
        ]);

        const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet);
        newSubscriptionsRows.push([
          newSubscriptionId,
          oldStudentID,
          oldSubscriptionType,
          teacherId,
          Utilities.formatDate(oldSubscriptionDate instanceof Date ? oldSubscriptionDate : new Date(oldSubscriptionDate), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          endDate,
          oldTotalAttendance || 0,
          oldRenewalStatus,
          Utilities.formatDate(oldLastPaymentDate instanceof Date ? oldLastPaymentDate : new Date(oldLastPaymentDate), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          oldAmount || 0,
          paidAmount,
          remainingAmount,
          0,
          ""
        ]);

        // 10. حجز المواعيد في شيت "المواعيد المتاحة للمعلمين" الجديد
        const bookedSlotsToUpdate = [];
        let bookingTypeForSlot = (oldRenewalStatus === "تجريبي") ? "تجريبي" : "عادي";

        // الميعاد الأول
        if (oldDay1 && oldTime1) {
            const timeSlotHeader1 = convertOldPlainTimeFormatToHeaderFormat(oldTime1);
            if (timeSlotHeader1) {
                bookedSlotsToUpdate.push({
                    teacherId: teacherId,
                    day: oldDay1,
                    timeSlotHeader: timeSlotHeader1,
                    studentId: oldStudentID,
                    bookingType: bookingTypeForSlot
                });
            } else {
                errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}): Invalid time format for M.1: '${String(oldTime1).trim()}'. Cannot book slot.`);
            }
        }
        // الميعاد الثاني
        if (oldDay2 && oldTime2) {
            const timeSlotHeader2 = convertOldPlainTimeFormatToHeaderFormat(oldTime2);
            if (timeSlotHeader2) {
                bookedSlotsToUpdate.push({
                    teacherId: teacherId,
                    day: oldDay2,
                    timeSlotHeader: timeSlotHeader2,
                    studentId: oldStudentID,
                    bookingType: bookingTypeForSlot
                });
            } else {
                errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}): Invalid time format for M.2: '${String(oldTime2).trim()}'. Cannot book slot.`);
            }
        }
        
        // تطبيق التحديثات على مصفوفة mutableTeachersAvailableSlotsData في الذاكرة
        bookedSlotsToUpdate.forEach(slot => {
            let teacherRowIndexInSlotsData = -1;
            for(let j = 0; j < mutableTeachersAvailableSlotsData.length; j++) {
                if (String(mutableTeachersAvailableSlotsData[j][0] || '').trim() === String(slot.teacherId).trim() &&
                    String(mutableTeachersAvailableSlotsData[j][1] || '').trim() === String(slot.day).trim()) {
                    teacherRowIndexInSlotsData = j;
                    break;
                }
            }

            if (teacherRowIndexInSlotsData === -1) {
                errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}) for slot ${slot.day} ${slot.timeSlotHeader}: Teacher/Day row not found in Teachers Available Slots sheet for Teacher ID: ${slot.teacherId}. Please ensure the Teachers Available Slots sheet is properly initialized.`);
                return;
            }

            const colIndex = timeSlotHeaderToColIndexMap.get(slot.timeSlotHeader);
            if (colIndex === undefined) {
                errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}) for slot ${slot.day} ${slot.timeSlotHeader}: Time slot header '${slot.timeSlotHeader}' not found in Teachers Available Slots sheet.`);
                return;
            }

            const currentCellValue = String(mutableTeachersAvailableSlotsData[teacherRowIndexInSlotsData][colIndex] || '').trim();

            if (currentCellValue === slot.studentId) {
                Logger.log(`Info: Slot ${slot.day} ${slot.timeSlotHeader} already booked by current student ${slot.studentId}. Skipping booking.`);
                return;
            }
            
            const isBookedByAnotherStudentOrOldValue = currentCellValue.startsWith("STD") || currentCellValue.startsWith("TRL") || currentCellValue.startsWith("p ");


            if (isBookedByAnotherStudentOrOldValue || (currentCellValue !== '' && currentCellValue !== slot.timeSlotHeader && !isTimeSlotHeaderFormat(currentCellValue))) {
                errors.push(`Error in row ${i + 1} (Student: ${oldStudentName}) for slot ${slot.day} ${slot.timeSlotHeader}: Slot already booked by '${currentCellValue}'.`);
                Logger.log(`Failed to book slot for ${oldStudentName}: Slot ${slot.day} ${slot.timeSlotHeader} already booked by '${currentCellValue}'.`);
            } else {
                mutableTeachersAvailableSlotsData[teacherRowIndexInSlotsData][colIndex] = slot.studentId;
                Logger.log(`تم حجز الميعاد ${slot.day} ${slot.timeSlotHeader} للطالب ${slot.studentId}.`);
            }
        });

        migratedCount++;

      } catch (e) {
        errors.push(`Error processing row ${i + 1} (Student: ${oldRow[0] || 'N/A'}): ${e.message}`);
        failedCount++;
        Logger.log(`FAILED Migration for row ${i + 1}: ${e.message}`);
      }
    }

    if (newStudentsRows.length > 0) {
        studentsSheet.getRange(studentsSheet.getLastRow() + 1, 1, newStudentsRows.length, newStudentsRows[0].length).setValues(newStudentsRows);
        Logger.log(`تم إضافة ${newStudentsRows.length} طالب جديد بالجملة.`);
    }

    if (newSubscriptionsRows.length > 0) {
        subscriptionsSheet.getRange(subscriptionsSheet.getLastRow() + 1, 1, newSubscriptionsRows.length, newSubscriptionsRows[0].length).setValues(newSubscriptionsRows);
        Logger.log(`تم إضافة ${newSubscriptionsRows.length} اشتراك جديد بالجملة.`);
    }

    teachersAvailableSlotsSheet.getRange(1, 1, mutableTeachersAvailableSlotsData.length, mutableTeachersAvailableSlotsData[0].length).setValues(mutableTeachersAvailableSlotsData);
    Logger.log(`تم تحديث ${migratedCount} طالب في مواعيد المعلمين بالجملة.`);


    const summary = `تمت عملية النقل بنجاح. تم نقل ${migratedCount} طالب. فشل نقل ${failedCount} طالب.`;
    Logger.log(summary);
    if (errors.length > 0) {
        Logger.log("تفاصيل الأخطاء:");
        errors.forEach(err => Logger.log(err));
    }
    return { success: summary, errors: errors };

  } catch (e) {
    Logger.log("خطأ عام في عملية النقل: " + e.message);
    return { error: `فشل عام في عملية النقل: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}






/**
 * دالة لتحديث عمود "Teacher ID" (العمود D) في شيت "الاشتراكات الحالية".
 * تقوم بجلب معرف المعلم من شيت "مواعيد الطلبة" ثم التحقق من شيت "المعلمين".
 *
 * المنطق:
 * 1. تمر على كل سطر في شيت "الاشتراكات الحالية".
 * 2. لكل "Student ID" (العمود B)، تبحث عنه في شيت "مواعيد الطلبة" (العمود A) في الملف الخارجي.
 * 3. من "مواعيد الطلبة"، تحصل على "اسم المعلم" (العمود E).
 * 4. باستخدام "اسم المعلم"، تبحث عنه في شيت "المعلمين" (العمود B) في الملف الخارجي للحصول على "Teacher ID" (العمود A).
 * 5. تعود إلى شيت "الاشتراكات الحالية" وتضع "Teacher ID" في العمود D لنفس السطر.
 *
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateSubscriptionsTeacherId() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  
  // تعريفات الشيتات (تأكد من مطابقة هذه الأسماء تماماً لأسماء الشيتات الفعلية لديك)
  const SUBSCRIPTIONS_SHEET_NAME = "الاشتراكات الحالية";
  const STUDENT_SCHEDULES_SHEET_NAME = "مواعيد الطلبة";
  const TEACHERS_SHEET_NAME = "المعلمين";

  // ***** الجديد: ID الملف الخارجي الذي يحتوي على مواعيد الطلبة والمعلمين *****
  const EXTERNAL_SHEET_FILE_ID = "1XWRFKp-LM7detp42X4bqjhVgkaJT00of6FWvtx8sYL0"; 
  // ************************************************************************

  const subscriptionsSheet = SPREADSHEET.getSheetByName(SUBSCRIPTIONS_SHEET_NAME);
  
  let externalSpreadsheet;
  try {
    externalSpreadsheet = SpreadsheetApp.openById(EXTERNAL_SHEET_FILE_ID);
  } catch (e) {
    Logger.log(`خطأ: تعذر فتح ملف Google Sheet الخارجي بالـ ID: ${EXTERNAL_SHEET_FILE_ID}. الخطأ: ${e.message}`);
    return { error: `تعذر فتح ملف المواعيد والمعلمين الخارجي. يرجى التأكد من الـ ID والأذونات.` };
  }

  const studentSchedulesSheet = externalSpreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
  const teachersSheet = externalSpreadsheet.getSheetByName(TEACHERS_SHEET_NAME);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // انتظر 30 ثانية للحصول على القفل

    // التحقق من وجود جميع الشيتات المطلوبة
    if (!subscriptionsSheet) throw new Error(`شيت '${SUBSCRIPTIONS_SHEET_NAME}' غير موجود في الملف الحالي.`);
    if (!studentSchedulesSheet) throw new Error(`شيت '${STUDENT_SCHEDULES_SHEET_NAME}' غير موجود في الملف الخارجي.`);
    if (!teachersSheet) throw new Error(`شيت '${TEACHERS_SHEET_NAME}' غير موجود في الملف الخارجي.`);

    // ------------------------------------------------------------------------
    // 1. قراءة جميع البيانات من الشيتات لتقليل عدد عمليات القراءة
    // ------------------------------------------------------------------------
    const subscriptionsData = subscriptionsSheet.getDataRange().getValues(); // بيانات الاشتراكات الحالية (في الملف الحالي)
    const studentSchedulesData = studentSchedulesSheet.getDataRange().getValues(); // بيانات مواعيد الطلبة (في الملف الخارجي)
    const teachersData = teachersSheet.getDataRange().getValues(); // بيانات المعلمين (في الملف الخارجي)

    if (subscriptionsData.length < 2) return { success: "لا توجد بيانات في شيت الاشتراكات الحالية لتحديثها." };
    if (studentSchedulesData.length < 2) return { error: "لا توجد بيانات في شيت مواعيد الطلبة لربطها." };
    if (teachersData.length < 2) return { error: "لا توجد بيانات في شيت المعلمين." };

    // ------------------------------------------------------------------------
    // 2. بناء خرائط مساعدة لسرعة البحث
    // ------------------------------------------------------------------------

    // خريطة: Teacher Name -> Teacher ID من الملف الخارجي
    // العمود A (مؤشر 0): Teacher ID
    // العمود B (مؤشر 1): اسم المعلم
    const teacherNameToIdMap = new Map();
    for (let i = 1; i < teachersData.length; i++) {
      const teacherId = String(teachersData[i][0] || '').trim();
      const teacherName = String(teachersData[i][1] || '').trim();
      if (teacherName) {
        teacherNameToIdMap.set(teacherName, teacherId);
      }
    }
    Logger.log(`تم بناء خريطة المعلمين من الملف الخارجي. عدد المعلمين: ${teacherNameToIdMap.size}`);

    // خريطة: Student ID -> Teacher Name من شيت "مواعيد الطلبة" في الملف الخارجي
    // العمود A (مؤشر 0): Student ID
    // العمود E (مؤشر 4): اسم المعلم
    // بما أن الطالب قد يكون له عدة مواعيد مع نفس المعلم، سنأخذ أول اسم معلم نجده له.
    const studentIdToTeacherNameFromSchedulesMap = new Map();
    for (let i = 1; i < studentSchedulesData.length; i++) {
      const studentId = String(studentSchedulesData[i][0] || '').trim(); // العمود A في مواعيد الطلبة
      const teacherNameFromSchedule = String(studentSchedulesData[i][4] || '').trim(); // العمود E في مواعيد الطلبة (اسم المعلم)
      if (studentId && teacherNameFromSchedule && !studentIdToTeacherNameFromSchedulesMap.has(studentId)) {
        studentIdToTeacherNameFromSchedulesMap.set(studentId, teacherNameFromSchedule);
      }
    }
    Logger.log(`تم بناء خريطة الطلاب إلى أسماء المعلمين من مواعيد الطلبة (الملف الخارجي). عدد الطلاب المربوطين: ${studentIdToTeacherNameFromSchedulesMap.size}`);

    // ------------------------------------------------------------------------
    // 3. المرور على شيت "الاشتراكات الحالية" وتحديث Teacher ID
    // ------------------------------------------------------------------------
    const updates = []; // لتخزين التحديثات التي ستطبق دفعة واحدة

    // العمود D هو العمود رقم 4 (مؤشر 3)
    const TEACHER_ID_COL_INDEX_IN_SUBSCRIPTIONS = 3; 

    for (let i = 1; i < subscriptionsData.length; i++) { // نبدأ من الصف الثاني (مؤشر 1) لتخطي العناوين
      const row = subscriptionsData[i];
      const studentId = String(row[1] || '').trim(); // العمود B (مؤشر 1): Student ID في الاشتراكات
      const currentTeacherIdInSubscription = String(row[TEACHER_ID_COL_INDEX_IN_SUBSCRIPTIONS] || '').trim(); // العمود D الحالي

      if (!studentId) {
        Logger.log(`تخطي الصف ${i + 1} في الاشتراكات: Student ID فارغ.`);
        continue;
      }

      // أ. البحث عن اسم المعلم للطالب من شيت "مواعيد الطلبة" (من الخريطة)
      const teacherNameForStudent = studentIdToTeacherNameFromSchedulesMap.get(studentId);

      if (!teacherNameForStudent) {
        // Logger.log(`لم يتم العثور على اسم معلم للطالب ${studentId} في شيت مواعيد الطلبة.`);
        continue; // تخطي هذا الطالب إذا لم نجد له معلماً في المواعيد
      }

      // ب. البحث عن Teacher ID من اسم المعلم (من الخريطة)
      const newTeacherId = teacherNameToIdMap.get(teacherNameForStudent);

      if (!newTeacherId) {
        Logger.log(`خطأ: لم يتم العثور على Teacher ID لاسم المعلم "${teacherNameForStudent}" (للطالب ${studentId}) في شيت المعلمين.`);
        continue; // تخطي هذا الطالب إذا لم نجد له Teacher ID
      }

      // ج. التحقق مما إذا كان التحديث مطلوباً
      if (newTeacherId !== currentTeacherIdInSubscription) {
        updates.push({
          rowIndex: i + 1, // رقم الصف في الشيت (1-based)
          colIndex: TEACHER_ID_COL_INDEX_IN_SUBSCRIPTIONS + 1, // رقم العمود في الشيت (1-based)
          value: newTeacherId
        });
        Logger.log(`تحديث Teacher ID للطالب ${studentId}: من "${currentTeacherIdInSubscription}" إلى "${newTeacherId}".`);
      }
    }

    // ------------------------------------------------------------------------
    // 4. تطبيق التحديثات دفعة واحدة على شيت "الاشتراكات الحالية"
    // ------------------------------------------------------------------------
    if (updates.length > 0) {
      updates.forEach(update => {
        subscriptionsSheet.getRange(update.rowIndex, update.colIndex).setValue(update.value);
      });
      Logger.log(`تم تطبيق ${updates.length} تحديث على عمود Teacher ID في شيت الاشتراكات الحالية.`);
      return { success: `تم تحديث ${updates.length} معرف معلم في شيت الاشتراكات الحالية بنجاح.` };
    } else {
      Logger.log("لا توجد تحديثات مطلوبة لعمود Teacher ID في شيت الاشتراكات الحالية.");
      return { success: "لا توجد تحديثات مطلوبة لعمود Teacher ID في شيت الاشتراكات الحالية." };
    }

  } catch (e) {
    Logger.log(`خطأ في دالة updateSubscriptionsTeacherId: ${e.message}`);
    return { error: `حدث خطأ أثناء تحديث Teacher ID في الاشتراكات: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}




function convertStudentIdToSubscriptionId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const slotsSheet = ss.getSheetByName("المواعيد المتاحة للمعلمين");
  const subscriptionsSheet = ss.getSheetByName("الاشتراكات الحالية");

  if (!slotsSheet || !subscriptionsSheet) {
    Logger.log("Sheet not found!");
    return;
  }

  // اقرأ بيانات الاشتراكات كـ مصفوفة
  const subsData = subscriptionsSheet.getDataRange().getValues();
  const subsHeaders = subsData[0];
  const subsRows = subsData.slice(1);

  const studentIdIndex = subsHeaders.indexOf("Student ID");
  const subscriptionIdIndex = subsHeaders.indexOf("Subscription ID");

  // أنشئ Map للربط بين Student ID و Subscription ID (أول اشتراك فقط)
  const studentToSubMap = {};
  for (let row of subsRows) {
    const studentId = row[studentIdIndex];
    const subId = row[subscriptionIdIndex];
    if (studentId && subId && !studentToSubMap[studentId]) {
      studentToSubMap[studentId] = subId;
    }
  }

  // اقرأ جدول المواعيد
  const slotsData = slotsSheet.getDataRange().getValues();

  // افترض أن الأعمدة الأولى 2 (المعلم واليوم) وباقي الأعمدة مواعيد
  const updatedData = slotsData.map((row, rowIndex) => {
    return row.map((cell, colIndex) => {
      // فقط من العمود الثالث فصاعدًا
      if (rowIndex === 0 || colIndex < 2) return cell;

      // لو القيمة Student ID موجود في Map → نستبدله بـ Subscription ID
      if (typeof cell === "string" && studentToSubMap[cell]) {
        return studentToSubMap[cell];
      }
      return cell;
    });
  });

  // كتابة البيانات المعدلة
  slotsSheet.getRange(1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);

  Logger.log("تم استبدال Student IDs بـ Subscription IDs بنجاح 🎉");
}
