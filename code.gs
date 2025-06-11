/**
 * هذا السكربت يحتوي على دوال لأتمتة شيت المشرف الرئيسي وشيت السجل التاريخي.
 */

// -----------------------------------------------------------------------------
// تعريفات ثابتة لأسماء الشيتات والأعمدة (تأكد من مطابقتها مع الشيت الفعلي)
// -----------------------------------------------------------------------------
const MASTER_SHEET_NAME = "بيانات الطلبة"; // اسم الشيت الرئيسي للطلاب (تأكد من هذا الاسم)
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
const HISTORY_MONTH_DATE_OFFSET = 1;     // الفهرس النسبي لعمود التاريخ داخل الثلاثية
const HISTORY_MONTH_AMOUNT_OFFSET = 2;   // الفهرس النسبي لعمود المبلغ داخل الثلاثية

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
// أعمدة الأرشيف ستكون مطابقة لأعمدة مواعيد الطلبة + عمودي الحذف
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
const ARCHIVE_DATE_DELETED_COL = 13; // M - **العمود الجديد لتاريخ الحذف**
const ARCHIVE_REASON_DELETED_COL = 14; // N - **العمود الجديد لسبب الحذف**
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




/**
 * هذه الدالة يتم تشغيلها تلقائياً عند فتح جدول البيانات.
 * تقوم بإنشاء قائمة مخصصة في شريط القوائم العلوي.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('لوحة تحكم المشرف') // اسم القائمة الرئيسية اللي هتظهر
      .addItem('فتح لوحة التحكم', 'openSupervisorPanel') // اسم العنصر في القائمة والدالة اللي هيستدعيها
      .addToUi();
}

/**
 * دالة لفتح واجهة المستخدم الرسومية (تطبيق الويب).
 * يتم تخصيصها لزر في القائمة الرئيسية في Google Sheet.
 */
function openSupervisorPanel() {
  const scriptUrl = ScriptApp.getService().getUrl(); // ده بيجيب رابط تطبيق الويب الحالي

  // ده هيفتح الواجهة في تبويبة جديدة في المتصفح
  const htmlOutput = HtmlService.createHtmlOutput(`<script>window.open('${scriptUrl}', '_blank'); google.script.host.close();</script>`)
      .setHeight(1)
      .setWidth(1)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME); // عشان تشتغل في بيئة Apps Script

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'جاري فتح لوحة التحكم...');
}




// **التعديل هنا في دالة doGet() فقط**
function doGet() {
  let teacherDataJson = "{}"; // قيمة افتراضية عشان نتجنب SyntaxError

  try {
    const data = getTeacherAndAvailableSchedules(); // الدالة دي المفروض بترجع JSON string
    if (data) {
      teacherDataJson = data; // لو البيانات رجعت سليمة
    }
  } catch (e) {
    Logger.log(`Error in doGet while fetching teacher data: ${e.message}`);
    // لو حصل خطأ في جلب البيانات، هتفضل teacherDataJson بـ "{}"
  }
  
  const htmlOutput = HtmlService.createTemplateFromFile('index');
  htmlOutput.initialTeacherDataJson = teacherDataJson; // تمرير البيانات كمتغير للقالب
  
  return htmlOutput.evaluate()
      .setTitle('إدارة أكاديمية رفاق')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}



/**
 * تجلب حالة التجديد لجميع الطلاب من شيت المشرف الرئيسي.
 * تُستخدم بواسطة واجهة الويب.
 * @returns {string} سلسلة JSON من مصفوفة كائنات الطلاب بحالتهم.
 */
function getStudentsRenewalStatus() {
  Logger.log("getStudentsRenewalStatus: Starting execution.");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المشرف الرئيسي باسم '${MASTER_SHEET_NAME}'.`);
    return JSON.stringify([]); // Return empty JSON string if sheet not found
  }

  const lastRow = masterSheet.getLastRow();
  if (lastRow < MASTER_FIRST_DATA_ROW) {
    Logger.log("getStudentsRenewalStatus: No data rows found. Returning empty array.");
    return JSON.stringify([]); // Return empty JSON string if no data rows
  }

  // نقرأ كل الأعمدة اللي ممكن نحتاجها في الواجهة
  const range = masterSheet.getRange(MASTER_FIRST_DATA_ROW, 1, lastRow - MASTER_FIRST_DATA_ROW + 1, masterSheet.getLastColumn());
  const values = range.getValues();
  Logger.log(`getStudentsRenewalStatus: Read ${values.length} rows from Master Sheet.`);

  const studentsData = [];
  values.forEach((row, index) => {
    const studentId = row[MASTER_STUDENT_ID_COL - 1];
    if (studentId && studentId.toString().trim() !== "") {
      try {
        studentsData.push({
          id: studentId,
          name: row[MASTER_STUDENT_NAME_COL - 1] || '',
          age: row[MASTER_STUDENT_AGE_COL - 1] || '', // **جديد: السن**
          number: row[MASTER_STUDENT_NUMBER_COL - 1] || '', // **جديد: رقم الهاتف**
          system: row[MASTER_SYSTEM_COL - 1] || '',
          attendance: row[MASTER_TOTAL_ATTENDANCE_COL - 1] || 0,
          renewalStatus: row[MASTER_RENEWAL_STATUS_COL - 1] || '',
          lastPaymentDate: row[MASTER_LAST_PAYMENT_DATE_COL - 1] ? formatDateOnly(new Date(row[MASTER_LAST_PAYMENT_DATE_COL - 1])) : '', // تنسيق التاريخ
          amount: row[MASTER_AMOUNT_COL - 1] || 0,
          row: MASTER_FIRST_DATA_ROW + index
        });
      } catch (e) {
        Logger.log(`خطأ في معالجة بيانات الطالب ${studentId} في الصف ${MASTER_FIRST_DATA_ROW + index}: ${e.message}`);
      }
    } else {
      Logger.log(`getStudentsRenewalStatus: Skipping empty student ID in row ${MASTER_FIRST_DATA_ROW + index}.`);
    }
  });

  Logger.log(`getStudentsRenewalStatus: Processed ${studentsData.length} valid students.`);
  Logger.log("Returning data as JSON string.");
  return JSON.stringify(studentsData);
}

/**
 * دالة جديدة تُستدعى من واجهة الويب لمعالجة التجديد.
 * إذا كان paidAmount > 0، فسيتم التعامل معه كدفعة جديدة.
 * إذا كان paidAmount == 0، فسيتم فقط تغيير حالة التجديد وتصفير الحصص.
 * @param {string} studentId معرف الطالب.
 * @param {number} row رقم الصف في شيت المشرف.
 * @param {number} paidAmount المبلغ المدفوع (0 إذا كان تجديد يدوي بدون مبلغ).
 */
function handleRenewalFromWeb(studentId, row, paidAmount) {
  Logger.log(`handleRenewalFromWeb: استدعاء لمعرف الطالب: ${studentId}, الصف: ${row}, المبلغ: ${paidAmount}`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    throw new Error(`لم يتم العثور على شيت المشرف الرئيسي باسم '${MASTER_SHEET_NAME}'.`);
  }

  const systemType = masterSheet.getRange(row, MASTER_SYSTEM_COL).getValue();
  const monthlyMaxClasses = getMonthlyMaxClasses(systemType);

  if (paidAmount > 0) {
    // إذا كان هناك مبلغ، تعامل معه كدفعة كاملة
    masterSheet.getRange(row, MASTER_RENEWAL_STATUS_COL).setValue("تم التجديد");
    masterSheet.getRange(row, MASTER_LAST_PAYMENT_DATE_COL).setValue(formatDateOnly(new Date()));

    const classesClearedForHistory = handleStudentRenewalClearance(studentId, monthlyMaxClasses);

    const studentName = masterSheet.getRange(row, MASTER_STUDENT_NAME_COL).getValue();
    const studentAge = masterSheet.getRange(row, MASTER_STUDENT_AGE_COL).getValue();
    const studentNumber = masterSheet.getRange(row, MASTER_STUDENT_NUMBER_COL).getValue();

    recordPaymentToHistory(studentId, studentName, studentAge, studentNumber, paidAmount, classesClearedForHistory);
    Logger.log(`تم تسجيل دفعة للطالب ${studentId} من الويب.`);
  } else {
    // إذا كان المبلغ 0، تعامل معه كتجديد يدوي يغير الحالة ويخصم الحصص فقط
    masterSheet.getRange(row, MASTER_RENEWAL_STATUS_COL).setValue("تم التجديد");
    masterSheet.getRange(row, MASTER_LAST_PAYMENT_DATE_COL).setValue(formatDateOnly(new Date()));
    handleStudentRenewalClearance(studentId, monthlyMaxClasses);
    Logger.log(`تم تجديد الطالب ${studentId} يدوياً من الويب (بدون مبلغ).`);
  }

  // تحديث الحضور وحالة التجديد بعد العملية
  const finalStudentAttendance = calculateStudentAttendance(studentId);
  updateSingleStudentAttendanceAndRenewalStatus(row, studentId, systemType, finalStudentAttendance);
  Logger.log(`تم تحديث شيت المشرف بعد عملية التجديد من الويب للطالب ${studentId}.`);
}





// -----------------------------------------------------------------------------
// دالة onEdit: تتفاعل مع التعديلات الفورية في شيت المشرف
// -----------------------------------------------------------------------------
function handleSheetEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedColumn = range.getColumn();
  const editedRow = range.getRow();
  const newValue = e.value;

  // سجلات تفصيلية جداً للتشخيص
  Logger.log("--- بدأ تشغيل handleSheetEdit ---");
  Logger.log(`اسم الشيت المعدل: ${sheet.getName()}`);
  Logger.log(`رقم الصف المعدل: ${editedRow}`);
  Logger.log(`رقم العمود المعدل: ${editedColumn}`);
  Logger.log(`القيمة الجديدة: ${newValue}`);
  Logger.log(`نوع القيمة الجديدة: ${typeof newValue}`); 

  // تحويل القيمة الجديدة إلى رقم هنا
  const parsedValue = parseFloat(newValue);
  Logger.log(`القيمة بعد التحويل (parsedValue): ${parsedValue}`);
  Logger.log(`نوع القيمة بعد التحويل (typeof parsedValue): ${typeof parsedValue}`);


  // تحقق من أن التعديل في الشيت الرئيسي وفي صفوف البيانات
  if (sheet.getName() === MASTER_SHEET_NAME && editedRow >= MASTER_FIRST_DATA_ROW) {
    Logger.log(`الشرط 1 (الشيت والصف) تحقق.`);

    const studentId = sheet.getRange(editedRow, MASTER_STUDENT_ID_COL).getValue();
    Logger.log(`معرف الطالب المحتمل: ${studentId}`);

    // تحقق إضافي: إذا كان معرف الطالب فارغًا، توقف
    if (!studentId || studentId.toString().trim() === "") {
      Logger.log(`handleSheetEdit: معرف الطالب في الصف ${editedRow} فارغ أو غير صالح. إلغاء العملية.`);
      return; // توقف العملية
    }
    Logger.log(`handleSheetEdit: معرف الطالب: ${studentId}`);


    const systemType = sheet.getRange(editedRow, MASTER_SYSTEM_COL).getValue();
    const monthlyMaxClasses = getMonthlyMaxClasses(systemType);

    // 2. الجزء الخاص بمسح علامات الحضور وتسجيل تاريخ آخر دفع عند التجديد (يتم عند تغيير عمود J يدوياً)
    if (editedColumn === MASTER_RENEWAL_STATUS_COL && newValue === "تم التجديد") {
      Logger.log(`الشرط 2 (تجديد الحالة) تحقق: تعديل يدوي لعمود حالة التجديد.`);
      
      // **** هنا يتم خصم الحصص ****
      // handleStudentRenewalClearance ستُرجع عدد الحصص المخصومة
      const classesCleared = handleStudentRenewalClearance(studentId, monthlyMaxClasses); 
      Logger.log(`handleSheetEdit: تم خصم ${classesCleared} حصة بعد التجديد اليدوي.`);
      
      sheet.getRange(editedRow, MASTER_LAST_PAYMENT_DATE_COL).setValue(formatDateOnly(new Date()));
      Browser.msgBox("تحديث و خصم الحصص", `تم تحديث الحضور للطالب ${studentId} وتم خصم ${classesCleared} حصة وتسجيل تاريخ الدفع بعد التجديد.`, Browser.Buttons.OK);
    }

    // 3. الجزء الجديد: تسجيل الدفع في السجل التاريخي عند تعديل عمود المبلغ (L)
    // هذا الشرط يتحقق من أن التعديل هو في عمود المبلغ وأن القيمة رقمية وصالحة وموجبة.
    if (editedColumn === MASTER_AMOUNT_COL && !isNaN(parsedValue) && parsedValue > 0) { 
      Logger.log(`الشرط 3 (تسجيل دفعة) تحقق: تعديل لعمود المبلغ.`);
      
      // **** هذا هو الجزء الذي يضمن الاتساق عند إدخال المبلغ ****
      
      // 1. تغيير حالة التجديد في عمود J إلى "تم التجديد" (تلقائياً)
      sheet.getRange(editedRow, MASTER_RENEWAL_STATUS_COL).setValue("تم التجديد");
      Logger.log(`تم تغيير حالة التجديد في العمود J إلى "تم التجديد" تلقائياً.`);

      // 2. تسجيل تاريخ الدفع في عمود K (تلقائياً)
      sheet.getRange(editedRow, MASTER_LAST_PAYMENT_DATE_COL).setValue(formatDateOnly(new Date()));
      Logger.log(`تم تسجيل تاريخ الدفع في العمود K تلقائياً.`);

      // 3. خصم الحصص من شيت المعلم (تلقائياً)
      // handleStudentRenewalClearance ستُرجع عدد الحصص المخصومة
      const classesClearedForHistory = handleStudentRenewalClearance(studentId, monthlyMaxClasses); 
      Logger.log(`تم استدعاء handleStudentRenewalClearance لخصم ${classesClearedForHistory} حصة تلقائياً.`);

      // 4. الحضور الذي سيتم إرساله إلى السجل التاريخي هو عدد الحصص التي تم خصمها
      const attendanceToRecordInHistory = classesClearedForHistory;
      Logger.log(`الحضور الذي سيتم إرساله إلى السجل التاريخي: ${attendanceToRecordInHistory}`);

      // 5. تسجيل الدفعة والحضور الجديد في السجل التاريخي
      const studentName = sheet.getRange(editedRow, MASTER_STUDENT_NAME_COL).getValue();
      const studentAge = sheet.getRange(editedRow, MASTER_STUDENT_AGE_COL).getValue();
      const studentNumber = sheet.getRange(editedRow, MASTER_STUDENT_NUMBER_COL).getValue();
      const paidAmount = parsedValue; // نمرر القيمة المحولة إلى رقم
      
      Logger.log(`استدعاء recordPaymentToHistory بمعرف: ${studentId}, مبلغ: ${paidAmount}, حضور: ${attendanceToRecordInHistory}`);
      recordPaymentToHistory(studentId, studentName, studentAge, studentNumber, paidAmount, attendanceToRecordInHistory);
      
      Browser.msgBox("تسجيل دفعة وتجديد", `تم تسجيل دفعة الطالب ${studentId} بالمبلغ ${paidAmount} وتجديد اشتراكه تلقائياً.`, Browser.Buttons.OK);

    } else if (editedColumn === MASTER_AMOUNT_COL && (isNaN(parsedValue) || parsedValue <= 0)) {
        Logger.log(`الشرط 3 (تسجيل دفعة) لم يتحقق: القيمة المدخلة غير رقمية أو غير موجبة. ` +
                   `editedColumn: ${editedColumn}, MASTER_AMOUNT_COL: ${MASTER_AMOUNT_COL}, ` +
                   `parsedValue: ${parsedValue}`);
    }


    // 4. الجزء الخاص بتحديث عمود "عدد الحصص الحاضرة" وحالة التجديد للصف الذي تم تعديله
    // هذا الجزء سيتم تشغيله بعد أي تعديل في الصف، لضمان تحديث عمود I و J بناءً على آخر حالة للحضور.
    // يتم حساب الحضور النهائي هنا بعد أي عمليات خصم تمت.
    const finalStudentAttendance = calculateStudentAttendance(studentId); 
    Logger.log(`تحديث شيت المشرف: الحضور النهائي (I) = ${finalStudentAttendance}`);
    // حالة التجديد (J) سيتم تحديثها أيضاً بناءً على الحضور النهائي، بغض النظر عن القيمة التي وضعها السكربت في الشرط 3.
    updateSingleStudentAttendanceAndRenewalStatus(editedRow, studentId, systemType, finalStudentAttendance);

    // رسائل تأكيد (يمكن تعديلها أو إزالتها لتفادياً للتكرار)
    // هذه الرسالة تظهر فقط إذا لم يكن التعديل في عمود التجديد أو عمود المبلغ (تفادياً للتكرار)
    if (editedColumn !== MASTER_RENEWAL_STATUS_COL && editedColumn !== MASTER_AMOUNT_COL) {
      Browser.msgBox("تحديث الحضور", `تم تحديث عدد الحصص وحالة التجديد للطالب ${studentId}.`, Browser.Buttons.OK);
    }
  } else {
      Logger.log(`handleSheetEdit: الشرط 1 (الشيت والصف) لم يتحقق. اسم الشيت: ${sheet.getName()}, الصف: ${editedRow}`);
  }
  Logger.log("--- انتهى تشغيل handleSheetEdit ---");
}



// -----------------------------------------------------------------------------
// دالة مساعدة جديدة: لحساب إجمالي حضور طالب واحد
// -----------------------------------------------------------------------------
function calculateStudentAttendance(studentId) {
  let totalStudentAttendance = 0;
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = spreadsheet.getSheetByName(TEACHERS_SHEET_NAME);
  if (!teachersSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'.`);
    throw new Error(`شيت المعلمين غير موجود: ${TEACHERS_SHEET_NAME}`);
  }
  const teacherValues = teachersSheet.getDataRange().getValues();

  const teacherUrls = [];
  for (let i = TEACHER_FIRST_DATA_ROW - 1; i < teacherValues.length; i++) {
    const teacherUrl = teacherValues[i][TEACHER_SHEET_URL_COL - 1];
    if (teacherUrl) {
      teacherUrls.push(teacherUrl);
    }
  }

  const allTeacherSheetsData = {};
  for (let urlCount = 0; urlCount < teacherUrls.length; urlCount++) {
    const teacherSheetUrl = teacherUrls[urlCount];
    try {
      const teacherSpreadsheet = SpreadsheetApp.openByUrl(teacherSheetUrl);
      const teacherSheet = teacherSpreadsheet.getSheets()[0];
      allTeacherSheetsData[teacherSheetUrl] = teacherSheet.getDataRange().getValues();
    } catch (e) {
      Logger.log(`خطأ في تهيئة بيانات المعلمين (calculateStudentAttendance) - الوصول أو معالجة شيت المعلم: ${teacherSheetUrl} - ${e.message}`);
    }
  }

  for (let j = 0; j < teacherUrls.length; j++) {
    const teacherSheetUrl = teacherUrls[j];
    const teacherSheetData = allTeacherSheetsData[teacherSheetUrl];

    if (teacherSheetData) {
      for (let k = 0; k < teacherSheetData.length; k++) {
        const rowData = teacherSheetData[k];

        // *** التعديل هنا: دمج التحقق من الصف مع التحقق من معرف الطالب ***
        // نتحقق أولاً من وجود الصف، ثم نتحقق من وجود معرف الطالب في الخلية
        if (rowData && rowData.length > (TEACHER_STUDENT_ID_COL_IN_TEACHER_SHEET - 1)) { // تأكد أن الصف طويل بما يكفي ليحتوي على عمود ID
          const currentTeacherStudentId = rowData[TEACHER_STUDENT_ID_COL_IN_TEACHER_SHEET - 1];
          
          // الآن نتحقق من currentTeacherStudentId قبل استخدام toString()
          if (currentTeacherStudentId !== undefined && currentTeacherStudentId !== null) { // هذا هو السطر 429
            if (currentTeacherStudentId.toString().trim() === studentId.toString().trim()) {
              for (let col = TEACHER_ATTENDANCE_COLS_START - 1; col < TEACHER_ATTENDANCE_COLS_END; col++) {
                // *** التحقق من وجود الخلية قبل الوصول إليها ***
                if (rowData.length > col && rowData[col] === true) {
                  totalStudentAttendance++;
                }
              }
              break;
            }
          }
        }
      }
    }
  }
  return totalStudentAttendance;
}



// -----------------------------------------------------------------------------
// دالة مساعدة: تقوم بحساب وتحديث الحضور وحالة التجديد لطالب واحد محدد
// (تم تعديلها لتستقبل totalStudentAttendance)
// -----------------------------------------------------------------------------
function updateSingleStudentAttendanceAndRenewalStatus(row, studentId, systemType, totalStudentAttendance) {
  Logger.log(`--- بدأ تشغيل updateSingleStudentAttendanceAndRenewalStatus للطالب ${studentId} ---`);
  Logger.log(`القيمة التي سيتم تحديثها في الحضور (I): ${totalStudentAttendance}`);
  Logger.log(`الصف المستهدف: ${row}, عمود الحضور (I): ${MASTER_TOTAL_ATTENDANCE_COL}`);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المشرف الرئيسي باسم '${MASTER_SHEET_NAME}'.`);
    return;
  }
  
  const monthlyMaxClasses = getMonthlyMaxClasses(systemType);
  let renewalStatus = "";

  if (monthlyMaxClasses > 0) {
      const renewalThreshold = monthlyMaxClasses; 
      if (totalStudentAttendance < renewalThreshold) { 
        renewalStatus = "تم التجديد"; 
      } else if (totalStudentAttendance === renewalThreshold) { 
        renewalStatus = "مطلوب التجديد"; 
      } else if (totalStudentAttendance > renewalThreshold) { 
        renewalStatus = "تجاوز الحد"; 
      } else if (totalStudentAttendance === 0 && studentId) { 
        renewalStatus = "لا يوجد حضور"; 
      } else {
        renewalStatus = "غير محدد"; 
      }
  } else {
      renewalStatus = ""; 
  }
  Logger.log(`حالة التجديد التي سيتم تعيينها (J): ${renewalStatus}`);
  
  try {
    masterSheet.getRange(row, MASTER_TOTAL_ATTENDANCE_COL).setValue(totalStudentAttendance);
    Logger.log(`تم تحديث عمود الحضور (I) للطالب ${studentId} بقيمة ${totalStudentAttendance}.`);
  } catch (e) {
    Logger.log(`خطأ في تحديث عمود الحضور (I) للطالب ${studentId}: ${e.message}`);
  }
  
  try {
    masterSheet.getRange(row, MASTER_RENEWAL_STATUS_COL).setValue(renewalStatus);
    Logger.log(`تم تحديث عمود حالة التجديد (J) للطالب ${studentId} بقيمة ${renewalStatus}.`);
  } catch (e) {
    Logger.log(`خطأ في تحديث عمود حالة التجديد (J) للطالب ${studentId}: ${e.message}`);
  }
  Logger.log(`--- انتهى تشغيل updateSingleStudentAttendanceAndRenewalStatus للطالب ${studentId} ---`);
}

// -----------------------------------------------------------------------------
// دالة مساعدة جديدة: لتحديد الحد الأقصى للحصص الشهرية بناءً على النظام
// -----------------------------------------------------------------------------
function getMonthlyMaxClasses(systemType) {
  if (typeof systemType !== 'string' || systemType.toString().trim() === "") {
    return 0; // إذا كان النظام فارغًا أو غير صالح، لا يوجد حد أقصى
  }
  const normalizedSystem = systemType.toString().trim().toLowerCase(); // تنظيف وتحويل لحالة الأحرف الصغيرة للمقارنة

  switch (normalizedSystem) {
    case "حلقة / اسبوع / 30 دقيقة":
    case "حلقة / اسبوع / 60 دقيقة":
      return 4;
    case "حلقتين / اسبوع / 30 دقيقة":
    case "حلقتين / اسبوع / 60 دقيقة":
    case "حلقتين / اسبوع / 45 دقيقة":
      return 8;
    case "3 حلقات / اسبوع / 30 دقيقة":
    case "3 حلقات / اسبوع / 45 دقيقة":
      return 12;
    case "4 حلقات / اسبوع / 30 دقيقة":
      return 16;
    case "6 حلقات / اسبوع / 30 دقيقة":
    case "يومياً": 
      return 24;
    default:
      // Logger.log(`نوع نظام غير معروف: ${systemType}`); // قم بالتعليق على هذا السطر أو حذفه
      return 0; 
  }
}

// -----------------------------------------------------------------------------
// دالة جديدة: لمعالجة التجديد (خصم عدد حصص الباقة من الحصص الأحدث)
// -----------------------------------------------------------------------------
// تم تعديل الدالة لترجع عدد الحصص التي تم خصمها
function handleStudentRenewalClearance(studentId, monthlyMaxClasses) {
  Logger.log(`--- بدأ تشغيل handleStudentRenewalClearance للطالب ${studentId}، الباقة ${monthlyMaxClasses} ---`);
  if (monthlyMaxClasses <= 0) { // لا يوجد باقة محددة، لا نفعل شيئًا
    Logger.log(`لم يتم تحديد باقة للطالب ${studentId} أو الباقة 0. لا توجد حصص لخصمها.`);
    return 0; // ارجع 0 حصة مخصومة
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const teachersSheet = spreadsheet.getSheetByName(TEACHERS_SHEET_NAME);
  if (!teachersSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'.`);
    return 0; // ارجع 0 حصة مخصومة
  }
  const teacherValues = teachersSheet.getDataRange().getValues();

  const teacherUrls = [];
  for (let i = TEACHER_FIRST_DATA_ROW - 1; i < teacherValues.length; i++) {
    const teacherUrl = teacherValues[i][TEACHER_SHEET_URL_COL - 1]; 
    if (teacherUrl) {
      teacherUrls.push(teacherUrl);
    }
  }

  let classesClearedCount = 0; // المتغير الذي سيخزن عدد الحصص المخصومة

  for (let i = 0; i < teacherUrls.length; i++) {
    const teacherSheetUrl = teacherUrls[i];
    try {
      const teacherSpreadsheet = SpreadsheetApp.openByUrl(teacherSheetUrl);
      const teacherSheet = teacherSpreadsheet.getSheets()[0]; 
      const teacherSheetData = teacherSheet.getDataRange().getValues();
      
      // البحث عن الطالب في شيت المعلم
      for (let j = 0; j < teacherSheetData.length; j++) {
        const currentTeacherStudentId = teacherSheetData[j][TEACHER_STUDENT_ID_COL_IN_TEACHER_SHEET - 1];
        
        if (currentTeacherStudentId !== undefined && currentTeacherStudentId !== null && currentTeacherStudentId.toString().trim() === studentId.toString().trim()) {
          const rowToModify = j + 1; // Apps Script يستخدم فهرسة 1 للصفوف
          const attendanceCells = []; // لتخزين نطاقات الخلايا التي تحتوي على علامات صح

          // نمر على أعمدة الحضور من F إلى Q ونجمع جميع الخلايا التي تحتوي على علامة صح
          for (let col = TEACHER_ATTENDANCE_COLS_START - 1; col < TEACHER_ATTENDANCE_COLS_END; col++) {
            if (teacherSheetData[j][col] === true) {
              attendanceCells.push(teacherSheet.getRange(rowToModify, col + 1)); // +1 لتحويل الفهرسة إلى رقم عمود
            }
          }
          Logger.log(`إجمالي الحصص المعلمة للطالب ${studentId} في هذا الشيت: ${attendanceCells.length}`);
          
          // **** المنطق الجديد لخصم الحصص عند التجديد ****
          let classesToClear = 0;
          if (attendanceCells.length > monthlyMaxClasses) {
              classesToClear = monthlyMaxClasses; // إذا تجاوز، خصم عدد حصص الباقة من الأحدث
          } else {
              classesToClear = attendanceCells.length; // إذا لم يتجاوز، خصم كل الحصص الحالية (تصفير)
          }
          Logger.log(`عدد الحصص المطلوب خصمها: ${classesToClear}`);

          classesClearedCount += classesToClear; // أضف إلى الإجمالي المخصوم

          if (classesToClear > 0) {
              // تحديد الخلايا التي سيتم مسحها (آخر 'classesToClear' حصص)
              const cellsToClear = attendanceCells.slice(-classesToClear);

              // الآن نقوم بمسح الحصص التي تم تحديدها فقط
              if (cellsToClear.length > 0) {
                cellsToClear.forEach(cellRange => {
                  cellRange.setValue(false); // تغيير قيمة الـ CheckBox إلى FALSE
                });
                Logger.log(`تم تغيير قيمة الخلايا في شيت المعلم.`);
                Logger.log(`تم خصم ${cellsToClear.length} حصة للطالب ${studentId} في شيت المعلم: ${teacherSheetUrl} بعد التجديد.`);
              }
          } else {
            Logger.log(`لا توجد حصص لخصمها للطالب ${studentId} في شيت المعلم: ${teacherSheetUrl} (الحصص الحالية صفر).`);
          }
          // ********************************************

          break; // الطالب وجد، لا حاجة للبحث أكثر في هذا الشيت
        }
      }
    } catch (error) {
      Logger.log(`خطأ في معالجة التجديد للطالب ${studentId} في شيت المعلم ${teacherSheetUrl}: ${error.message}`);
    }
  }
  Logger.log(`--- انتهى تشغيل handleStudentRenewalClearance للطالب ${studentId} ---`);
  return classesClearedCount; // ارجع العدد الإجمالي للحصص المخصومة
}

// -----------------------------------------------------------------------------
// دالة مساعدة: لتحويل رقم الشهر إلى اسمه العربي
// -----------------------------------------------------------------------------
function getMonthName(monthNumber) { // 0-indexed month (0 for Jan, 11 for Dec)
  return MONTH_NAMES_AR[monthNumber];
}






/**
 * دالة مساعدة لتنسيق كائن التاريخ إلى صيغة "YYYY-MM-DD" أو "DD/MM/YYYY".
 * @param {Date} dateObj كائن التاريخ المراد تنسيقه.
 * @param {string} formatType نوع التنسيق المطلوب ('YYYY-MM-DD' أو 'DD/MM/YYYY').
 * @returns {string} التاريخ منسقاً.
 */
function formatDateOnly(dateObj, formatType = 'DD/MM/YYYY') {
  if (!dateObj || !(dateObj instanceof Date)) {
    return ''; // ارجع سلسلة فارغة إذا لم يكن تاريخًا صالحًا
  }

  const year = dateObj.getFullYear();
  const month = (dateObj.getMonth() + 1).toString().padStart(2, '0'); // الشهر 0-indexed
  const day = dateObj.getDate().toString().padStart(2, '0');

  if (formatType === 'YYYY-MM-DD') {
    return `${year}-${month}-${day}`;
  } else { // الافتراضي DD/MM/YYYY
    return `${day}/${month}/${year}`;
  }
}






// -----------------------------------------------------------------------------
// دالة مساعدة: للبحث عن عمود باسمه في صف الرؤوس
// -----------------------------------------------------------------------------
// هذه الدالة لم تعد تستخدم بنفس الطريقة بعد التعديلات الأخيرة،
// لكنها لا تزال موجودة في الكود.
// الوظائف الجديدة تستخدم البحث المباشر في المصفوفات.
function findColumnIndexByNameInRow(sheet, rowNumber, columnName) {
  const headers = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === columnName.trim()) {
      return i + 1; // Apps Script columns are 1-indexed
    }
  }
  return -1; // Column not found
}


// -----------------------------------------------------------------------------
// دالة جديدة: لتسجيل الدفعات في شيت "السجل التاريخي" (معدلة للهيكل الجديد)
// -----------------------------------------------------------------------------
// **** هنا التعديل: تغيير اسم المعامل من totalStudentAttendance إلى attendanceToRecordInHistory ****
function recordPaymentToHistory(studentId, studentName, studentAge, studentNumber, paidAmount, attendanceToRecordInHistory) { 
  Logger.log("--- بدأ تشغيل recordPaymentToHistory ---");
  Logger.log(`بيانات الدفعة: معرف الطالب=${studentId}, المبلغ=${paidAmount}, الحضور=${attendanceToRecordInHistory}, الشهر=${getMonthName(new Date().getMonth())}`);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = spreadsheet.getSheetByName(HISTORY_SHEET_NAME);
  if (!historySheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت السجل التاريخي باسم '${HISTORY_SHEET_NAME}'. إنهاء العملية.`);
    Browser.msgBox("خطأ", `لم يتم العثور على شيت السجل التاريخي باسم '${HISTORY_SHEET_NAME}'. يرجى التحقق من الاسم.`, Browser.Buttons.OK);
    return;
  }
  Logger.log(`تم العثور على شيت السجل التاريخي: ${HISTORY_SHEET_NAME}`);


  const today = formatDateOnly(new Date());
  const currentMonthName = getMonthName(today.getMonth());
  const paymentDate = today;

  const historyHeadersRow1 = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
  Logger.log(`رؤوس الصف 1: ${historyHeadersRow1}`);


  // 1. تحديد أعمدة الشهر الحالي في السجل التاريخي (الثلاثية)
  let monthStartColIndex = -1; // عمود بداية الشهر المدمج (مثلاً E لمارس)
  for(let i=0; i < historyHeadersRow1.length; i++){
    if(historyHeadersRow1[i].toString().trim() === currentMonthName){
      monthStartColIndex = i + 1; // 1-indexed column number
      break;
    }
  }
  Logger.log(`تم تحديد الشهر المدمج: ${currentMonthName} في العمود: ${monthStartColIndex}`);

  // تم إزالة منطق الإدراج التلقائي للأعمدة الشهرية بناءً على طلبك.
  // يجب أن تقوم بإضافة الأعمدة (عدد الحصص، التاريخ، المبلغ) يدويًا لكل شهر جديد.
  if (monthStartColIndex === -1) {
    Logger.log(`تحذير: أعمدة الشهر الحالي '${currentMonthName}' غير موجودة. لن يتم تسجيل الدفعة الشهرية حتى يتم إضافة الأعمدة يدويًا.`);
    Browser.msgBox("تنبيه", `أعمدة شهر ${currentMonthName} غير موجودة في "السجل التاريخي". يرجى إضافتها يدويًا لتسجيل الدفعات الشهرية.`, Browser.Buttons.OK);
    return; // توقف عن معالجة الدفعة إذا لم يتم العثور على أعمدة الشهر
  } else {
    Logger.log(`أعمدة الشهر الحالي '${currentMonthName}' موجودة بالفعل. بداية العمود=${monthStartColIndex}`);
  }

  // الآن لدينا monthStartColIndex لعمود بداية ثلاثية الشهر
  const currentMonthAttendanceColIndex = monthStartColIndex + HISTORY_MONTH_ATTENDANCE_OFFSET;
  const currentMonthDateColIndex = monthStartColIndex + HISTORY_MONTH_DATE_OFFSET;
  const currentMonthAmountColIndex = monthStartColIndex + HISTORY_MONTH_AMOUNT_OFFSET;
  Logger.log(`فهارس الأعمدة النهائية للشهر: حضور=${currentMonthAttendanceColIndex}, تاريخ=${currentMonthDateColIndex}, مبلغ=${currentMonthAmountColIndex}`);


  // 2. البحث عن الطالب في السجل التاريخي (من الصف الثالث فصاعدًا)
  const lastHistoryRow = historySheet.getLastRow();
  const numRowsToRead = lastHistoryRow - HISTORY_FIRST_DATA_ROW + 1;
  const historyDataRange = (numRowsToRead > 0) ? historySheet.getRange(HISTORY_FIRST_DATA_ROW, 1, numRowsToRead, historySheet.getLastColumn()) : null;
  const historyData = historyDataRange ? historyDataRange.getValues() : [];
  Logger.log(`تم قراءة بيانات السجل التاريخي من الصف ${HISTORY_FIRST_DATA_ROW} حتى ${lastHistoryRow}. عدد الصفوف المقروءة: ${historyData.length}`);

  let studentFoundRow = -1; // سيكون فهرس الصف في البيانات المحملة (0-indexed)
  for (let i = 0; i < historyData.length; i++) {
    const currentStudentIdInHistory = historyData[i][HISTORY_STUDENT_ID_COL - 1];
    if (currentStudentIdInHistory !== undefined && currentStudentIdInHistory !== null && 
        currentStudentIdInHistory.toString().trim() === studentId.toString().trim()) { 
      studentFoundRow = i; 
      break;
    }
  }
  Logger.log(`نتيجة البحث عن الطالب: studentFoundRow (index in data) = ${studentFoundRow}`);


  let targetRowInSheet; // رقم الصف الفعلي في الشيت (1-indexed)

  if (studentFoundRow === -1) {
    // الطالب غير موجود، أضف صفًا جديدًا في نهاية الشيت
    targetRowInSheet = historySheet.getLastRow() + 1;
    const numCols = historySheet.getLastColumn();
    const newRowData = Array(numCols).fill(''); 
    
    newRowData[HISTORY_STUDENT_ID_COL - 1] = studentId;
    newRowData[HISTORY_STUDENT_NAME_COL - 1] = studentName;
    newRowData[HISTORY_STUDENT_AGE_COL - 1] = studentAge;
    newRowData[HISTORY_STUDENT_NUMBER_COL - 1] = studentNumber;

    historySheet.getRange(targetRowInSheet, 1, 1, numCols).setValues([newRowData]);
    Logger.log(`الطالب غير موجود. تم إضافة صف جديد في: ${targetRowInSheet}`);
  } else {
    // الطالب موجود، استخدم الصف الحالي
    targetRowInSheet = HISTORY_FIRST_DATA_ROW + studentFoundRow;
    Logger.log(`الطالب موجود في الصف: ${targetRowInSheet}`);
  }

  // 3. تحديث بيانات الدفع الشهرية للطالب في صفه الخاص
  // نكتب في الأعمدة الثلاثة: عدد الحصص، التاريخ، المبلغ
  if(currentMonthAttendanceColIndex !== -1 && currentMonthDateColIndex !== -1 && currentMonthAmountColIndex !== -1){
    historySheet.getRange(targetRowInSheet, currentMonthAttendanceColIndex).setValue(attendanceToRecordInHistory); // تسجيل الحضور
    historySheet.getRange(targetRowInSheet, currentMonthDateColIndex).setValue(paymentDate); // تسجيل التاريخ
    historySheet.getRange(targetRowInSheet, currentMonthAmountColIndex).setValue(paidAmount); // تسجيل المبلغ
    Logger.log(`تم تحديث بيانات الدفع في الخلايا: ` + 
               `${historySheet.getRange(targetRowInSheet, currentMonthAttendanceColIndex).getA1Notation()}, ` +
               `${historySheet.getRange(targetRowInSheet, currentMonthDateColIndex).getA1Notation()}, ` +
               `${historySheet.getRange(targetRowInSheet, currentMonthAmountColIndex).getA1Notation()}`);
  } else {
    Logger.log(`تحذير: لا يمكن تحديث بيانات الدفع للشهر. فهارس الأعمدة غير صالحة.`);
  }

  // تم حذف الجزء الخاص بتحديث عمود "إجمالي المدفوع" (العمود T) أو وضع صيغة SUM فيه، بناءً على طلبك.
  // أنت ستدير هذا العمود يدويًا.

  Logger.log("--- انتهى تشغيل recordPaymentToHistory ---");
}

// -----------------------------------------------------------------------------
// دالة التحديث الشاملة: تقوم بتحديث جميع الطلاب (مع تحسين الأداء)
// -----------------------------------------------------------------------------
function updateAllStudentsAttendanceAndRenewalStatus() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المشرف الرئيسي باسم '${MASTER_SHEET_NAME}'.`);
    Browser.msgBox("خطأ", `لم يتم العثور على شيت الطلاب الرئيسي باسم '${MASTER_SHEET_NAME}'. يرجى التحقق من الاسم.`, Browser.Buttons.OK);
    return;
  }
  const masterValues = masterSheet.getDataRange().getValues();

  // جلب بيانات المعلمين مرة واحدة وتخزينها (تحسين الأداء)
  const teachersSheet = spreadsheet.getSheetByName(TEACHERS_SHEET_NAME);
  if (!teachersSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'.`);
    Browser.msgBox("خطأ", `لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'. يرجى التحقق من الاسم.`, Browser.Buttons.OK);
    return;
  }
  const teacherValues = teachersSheet.getDataRange().getValues();

  const teacherUrls = [];
  for (let urlIndex = TEACHER_FIRST_DATA_ROW - 1; urlIndex < teacherValues.length; urlIndex++) {
    const teacherUrl = teacherValues[urlIndex][TEACHER_SHEET_URL_COL - 1]; 
    if (teacherUrl) {
      teacherUrls.push(teacherUrl);
    }
  }

  const allTeacherSheetsData = {}; 
  for (let urlCount = 0; urlCount < teacherUrls.length; urlCount++) {
    const teacherSheetUrl = teacherUrls[urlCount];
    try {
      const teacherSpreadsheet = SpreadsheetApp.openByUrl(teacherSheetUrl);
      const teacherSheet = teacherSpreadsheet.getSheets()[0]; 
      allTeacherSheetsData[teacherSheetUrl] = teacherSheet.getDataRange().getValues(); 
    } catch (e) {
      Logger.log(`خطأ في تهيئة بيانات المعلمين (updateAllStudentsAttendanceAndRenewalStatus) - الوصول أو معالجة شيت المعلم: ${teacherSheetUrl} - ${e.message}`);
    }
  }

  const updatedAttendance = [];
  const updatedRenewalStatus = [];

  // التكرار على كل طالب في شيت المشرف الرئيسي
  for (let i = MASTER_FIRST_DATA_ROW - 1; i < masterValues.length; i++) {
    const studentId = masterValues[i][MASTER_STUDENT_ID_COL - 1];
    const systemType = masterValues[i][MASTER_SYSTEM_COL - 1]; 

    if (!studentId || studentId.toString().trim() === "") {
      updatedAttendance.push([""]); 
      updatedRenewalStatus.push([""]);
      continue;
    }

    let totalStudentAttendance = 0; 
    
    // حساب إجمالي الحضور للطالب من جميع المعلمين
    for (let urlCount = 0; urlCount < teacherUrls.length; urlCount++) {
      const teacherSheetUrl = teacherUrls[urlCount];
      const teacherSheetData = allTeacherSheetsData[teacherSheetUrl]; 

      if (teacherSheetData) { 
        for (let k = 0; k < teacherSheetData.length; k++) {
          const currentTeacherStudentId = teacherSheetData[k][TEACHER_STUDENT_ID_COL_IN_TEACHER_SHEET - 1];
          if (currentTeacherStudentId !== undefined && currentTeacherStudentId !== null && currentTeacherStudentId.toString().trim() === studentId.toString().trim()) {
            let studentAttendanceInThisTeacherSheet = 0;
            for (let col = TEACHER_ATTENDANCE_COLS_START - 1; col < TEACHER_ATTENDANCE_COLS_END; col++) {
              if (teacherSheetData[k][col] === true) {
                studentAttendanceInThisTeacherSheet++;
              }
            }
            totalStudentAttendance += studentAttendanceInThisTeacherSheet;
            break; 
          }
        }
      }
    }

    // -----------------------------------------------------------------------------
    // تحديد الحد الأقصى للحصص بناءً على "النظام"
    const monthlyMaxClasses = getMonthlyMaxClasses(systemType);
    
    let renewalStatus = "";

    if (monthlyMaxClasses > 0) { // إذا كان هناك نظام صالح ومحدد
        const renewalThreshold = monthlyMaxClasses; 

        if (totalStudentAttendance < renewalThreshold) { 
          renewalStatus = "تم التجديد"; 
        } else if (totalStudentAttendance === renewalThreshold) { 
          renewalStatus = "مطلوب التجديد"; 
        } else if (totalStudentAttendance > renewalThreshold) { 
          renewalStatus = "تجاوز الحد"; 
        } else if (totalStudentAttendance === 0 && studentId) { 
          renewalStatus = "لا يوجد حضور"; 
        } else {
          renewalStatus = "غير محدد"; 
        }
    } else { // إذا كان monthlyMaxClasses === 0 (أي النظام فارغ أو غير صالح)
        renewalStatus = ""; 
    }

    updatedAttendance.push([totalStudentAttendance]);
    updatedRenewalStatus.push([renewalStatus]);
  }

  // تحديث شيت المشرف الرئيسي بالنتائج دفعة واحدة لتحسين الأداء
  if (updatedAttendance.length > 0) {
    const startRow = MASTER_FIRST_DATA_ROW;
    const numRows = updatedAttendance.length;

    masterSheet.getRange(startRow, MASTER_TOTAL_ATTENDANCE_COL, numRows, 1).setValues(updatedAttendance);
    masterSheet.getRange(startRow, MASTER_RENEWAL_STATUS_COL, numRows, 1).setValues(updatedRenewalStatus);
  }

  Browser.msgBox("اكتمل التحديث اليومي", "تم تحديث عدد الحصص وحالة التجديد لجميع الطلاب بنجاح.", Browser.Buttons.OK);
}

// -----------------------------------------------------------------------------
// دالة التحديث الشاملة: تقوم بتحديث جميع الطلاب (مع تحسين الأداء)
// -----------------------------------------------------------------------------
function updateAllStudentsAttendanceAndRenewalStatus() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المشرف الرئيسي باسم '${MASTER_SHEET_NAME}'.`);
    Browser.msgBox("خطأ", `لم يتم العثور على شيت الطلاب الرئيسي باسم '${MASTER_SHEET_NAME}'. يرجى التحقق من الاسم.`, Browser.Buttons.OK);
    return;
  }
  const masterValues = masterSheet.getDataRange().getValues();

  // جلب بيانات المعلمين مرة واحدة وتخزينها (تحسين الأداء)
  const teachersSheet = spreadsheet.getSheetByName(TEACHERS_SHEET_NAME);
  if (!teachersSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'.`);
    Browser.msgBox("خطأ", `لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'. يرجى التحقق من الاسم.`, Browser.Buttons.OK);
    return;
  }
  const teacherValues = teachersSheet.getDataRange().getValues();

  const teacherUrls = [];
  for (let urlIndex = TEACHER_FIRST_DATA_ROW - 1; urlIndex < teacherValues.length; urlIndex++) {
    const teacherUrl = teacherValues[urlIndex][TEACHER_SHEET_URL_COL - 1]; 
    if (teacherUrl) {
      teacherUrls.push(teacherUrl);
    }
  }

  const allTeacherSheetsData = {}; 
  for (let urlCount = 0; urlCount < teacherUrls.length; urlCount++) {
    const teacherSheetUrl = teacherUrls[urlCount];
    try {
      const teacherSpreadsheet = SpreadsheetApp.openByUrl(teacherSheetUrl);
      const teacherSheet = teacherSpreadsheet.getSheets()[0]; 
      allTeacherSheetsData[teacherSheetUrl] = teacherSheet.getDataRange().getValues(); 
    } catch (e) {
      Logger.log(`خطأ في تهيئة بيانات المعلمين (updateAllStudentsAttendanceAndRenewalStatus) - الوصول أو معالجة شيت المعلم: ${teacherSheetUrl} - ${e.message}`);
    }
  }

  const updatedAttendance = [];
  const updatedRenewalStatus = [];

  // التكرار على كل طالب في شيت المشرف الرئيسي
  for (let i = MASTER_FIRST_DATA_ROW - 1; i < masterValues.length; i++) {
    const studentId = masterValues[i][MASTER_STUDENT_ID_COL - 1];
    const systemType = masterValues[i][MASTER_SYSTEM_COL - 1]; 

    if (!studentId || studentId.toString().trim() === "") {
      updatedAttendance.push([""]); 
      updatedRenewalStatus.push([""]);
      continue;
    }

    let totalStudentAttendance = 0; 
    
    // حساب إجمالي الحضور للطالب من جميع المعلمين
    for (let urlCount = 0; urlCount < teacherUrls.length; urlCount++) {
      const teacherSheetUrl = teacherUrls[urlCount];
      const teacherSheetData = allTeacherSheetsData[teacherSheetUrl]; 

      if (teacherSheetData) { 
        for (let k = 0; k < teacherSheetData.length; k++) {
          const currentTeacherStudentId = teacherSheetData[k][TEACHER_STUDENT_ID_COL_IN_TEACHER_SHEET - 1];
          if (currentTeacherStudentId !== undefined && currentTeacherStudentId !== null && currentTeacherStudentId.toString().trim() === studentId.toString().trim()) {
            let studentAttendanceInThisTeacherSheet = 0;
            for (let col = TEACHER_ATTENDANCE_COLS_START - 1; col < TEACHER_ATTENDANCE_COLS_END; col++) {
              if (teacherSheetData[k][col] === true) {
                studentAttendanceInThisTeacherSheet++;
              }
            }
            totalStudentAttendance += studentAttendanceInThisTeacherSheet;
            break; 
          }
        }
      }
    }

    // -----------------------------------------------------------------------------
    // تحديد الحد الأقصى للحصص بناءً على "النظام"
    const monthlyMaxClasses = getMonthlyMaxClasses(systemType);
    
    let renewalStatus = "";

    if (monthlyMaxClasses > 0) { // إذا كان هناك نظام صالح ومحدد
        const renewalThreshold = monthlyMaxClasses; 

        if (totalStudentAttendance < renewalThreshold) { 
          renewalStatus = "تم التجديد"; 
        } else if (totalStudentAttendance === renewalThreshold) { 
          renewalStatus = "مطلوب التجديد"; 
        } else if (totalStudentAttendance > renewalThreshold) { 
          renewalStatus = "تجاوز الحد"; 
        } else if (totalStudentAttendance === 0 && studentId) { 
          renewalStatus = "لا يوجد حضور"; 
        } else {
          renewalStatus = "غير محدد"; 
        }
    } else { // إذا كان monthlyMaxClasses === 0 (أي النظام فارغ أو غير صالح)
        renewalStatus = ""; 
    }

    updatedAttendance.push([totalStudentAttendance]);
    updatedRenewalStatus.push([renewalStatus]);
  }

  // تحديث شيت المشرف الرئيسي بالنتائج دفعة واحدة لتحسين الأداء
  if (updatedAttendance.length > 0) {
    const startRow = MASTER_FIRST_DATA_ROW;
    const numRows = updatedAttendance.length;

    masterSheet.getRange(startRow, MASTER_TOTAL_ATTENDANCE_COL, numRows, 1).setValues(updatedAttendance);
    masterSheet.getRange(startRow, MASTER_RENEWAL_STATUS_COL, numRows, 1).setValues(updatedRenewalStatus);
  }

  Browser.msgBox("اكتمل التحديث اليومي", "تم تحديث عدد الحصص وحالة التجديد لجميع الطلاب بنجاح.", Browser.Buttons.OK);
}


// -----------------------------------------------------------------------------
// دالة مساعدة: لتحويل رقم العمود إلى حرف العمود (A, B, C...)
// يجب أن تكون هذه الدالة موجودة لتعمل Utilities.getColLetter
// -----------------------------------------------------------------------------
Utilities.getColLetter = function(col) {
  let letter = '';
  while (col > 0) {
    col--;
    letter = String.fromCharCode(65 + (col % 26)) + letter;
    col = Math.floor(col / 26);
  }
  return letter;
};





// ... (كل الكود القديم والدوال الموجودة زي ما هي) ...

// -----------------------------------------------------------------------------
// دوال جديدة لإدارة إضافة الطلاب والمواعيد
// -----------------------------------------------------------------------------

/**
 * تجلب أسماء جميع المعلمين والمواعيد المتاحة لكل معلم، وحالة المواعيد المحجوزة.
 * تُستخدم لملء القوائم المنسدلة في الواجهة.
 * @returns {string} JSON string containing { teacherNames: [], availableSchedules: {}, allSchedulesWithStatus: {} }
 */
function getTeacherAndAvailableSchedules() {
  Logger.log("getTeacherAndAvailableSchedules: Starting execution.");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let teacherNames = [];
  let availableSchedules = {};
  let allSchedulesWithStatus = {};

  try {
    // 1. جلب أسماء المعلمين الفريدة من شيت "المعلمين"
    const teachersSheet = spreadsheet.getSheetByName(TEACHERS_SHEET_NAME);
    if (!teachersSheet) {
      Logger.log(`خطأ: لم يتم العثور على شيت المعلمين باسم '${TEACHERS_SHEET_NAME}'.`);
      throw new Error(`شيت المعلمين غير موجود: ${TEACHERS_SHEET_NAME}`);
    }
    const allTeachersSheetValues = teachersSheet.getRange(TEACHER_FIRST_DATA_ROW, TEACHER_ID_COL, teachersSheet.getLastRow() - TEACHER_FIRST_DATA_ROW + 1, TEACHER_NAME_COL).getValues();
    const uniqueTeacherNamesSet = new Set();
    for (let i = 0; i < allTeachersSheetValues.length; i++) {
        const row = allTeachersSheetValues[i];
        const name = row[TEACHER_NAME_COL - 1];
        if (name && name.toString().trim() !== "") {
            uniqueTeacherNamesSet.add(name.toString().trim());
        }
    }
    teacherNames = Array.from(uniqueTeacherNamesSet);
    Logger.log(`تم جلب ${teacherNames.length} اسم معلم فريد من شيت المعلمين: ${JSON.stringify(teacherNames)}`);


    // 2. جلب المواعيد المتاحة من شيت "المواعيد المتاحة للمعلمين"
    const availableSchedulesSheet = spreadsheet.getSheetByName(AVAILABLE_SCHEDULES_SHEET_NAME);
    if (!availableSchedulesSheet) {
      Logger.log(`خطأ: لم يتم العثور على شيت المواعيد المتاحة للمعلمين باسم '${AVAILABLE_SCHEDULES_SHEET_NAME}'.`);
      throw new Error(`شيت المواعيد المتاحة للمعلمين غير موجود: ${AVAILABLE_SCHEDULES_SHEET_NAME}`);
    }

    const allScheduleData = availableSchedulesSheet.getDataRange().getDisplayValues();
    Logger.log(`عدد الصفوف المقروءة من المواعيد المتاحة: ${allScheduleData.length}`);

    if (allScheduleData.length < AVAILABLE_SCHEDULE_FIRST_DATA_ROW) {
        Logger.log("شيت المواعيد المتاحة للمعلمين لا يحتوي على بيانات كافية.");
        return JSON.stringify({ teacherNames: teacherNames, availableSchedules: {}, allSchedulesWithStatus: {} });
    }

    // --- **جمع كل فهارس أعمدة المواعيد الموجودة فعلاً** ---
    let actualTimeColIndices = new Set();
    // نمر على كل صفوف البيانات (من الصف الأول للبيانات)
    for (let i = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; i < allScheduleData.length; i++) {
        const row = allScheduleData[i];
        // نمر على الأعمدة من أول عمود المواعيد (D)
        for (let col = AVAILABLE_SCHEDULE_TIMES_START_COL - 1; col < row.length; col++) {
            const timeValue = row[col];
            if (timeValue && timeValue.toString().trim() !== "") {
                // نضيف فهرس العمود (0-indexed) فقط
                actualTimeColIndices.add(col);
            }
        }
    }
    const sortedActualTimeColIndices = Array.from(actualTimeColIndices).sort((a,b) => a - b);
    Logger.log("Actual Time Column Indices found (sorted):", JSON.stringify(sortedActualTimeColIndices));
    // -------------------------------------------------------------

    for (let i = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; i < allScheduleData.length; i++) {
      const row = allScheduleData[i];
      const teacherNameInSchedule = row[AVAILABLE_SCHEDULE_TEACHER_NAME_COL - 1];
      const day = row[AVAILABLE_SCHEDULE_DAY_COL - 1];

      if (teacherNameInSchedule && day) {
          const currentTeacherName = teacherNameInSchedule.toString().trim();
          const currentDay = day.toString().trim();

          // تهيئة الكائنات للمعلم واليوم لو مش موجودين
          if (!availableSchedules[currentTeacherName]) {
              availableSchedules[currentTeacherName] = {};
              allSchedulesWithStatus[currentTeacherName] = {};
          }
          if (!availableSchedules[currentTeacherName][currentDay]) {
              availableSchedules[currentTeacherName][currentDay] = [];
              allSchedulesWithStatus[currentTeacherName][currentDay] = [];
          }

          // نمر على فهارس الأعمدة اللي فيها مواعيد فعلية (sortedActualTimeColIndices)
          for (let k = 0; k < sortedActualTimeColIndices.length; k++) {
              const timeColIndex = sortedActualTimeColIndices[k];
              const actualTimeValueInRow = row[timeColIndex]; // القيمة الفعلية في الخلية

              if (actualTimeValueInRow && actualTimeValueInRow.toString().trim() !== "") {
                  const trimmedTime = actualTimeValueInRow.toString().trim();
                  // ***** لا يوجد تحويل هنا، نستخدم النص كما هو *****
                  availableSchedules[currentTeacherName][currentDay].push(trimmedTime);
                  allSchedulesWithStatus[currentTeacherName][currentDay].push({
                      timeSlot: trimmedTime, // نخزن النص كما هو
                      isBooked: false,
                      bookedBy: null,
                      colIndex: timeColIndex + 1 // نضيف فهرس العمود (1-indexed)
                  });
              }
          }
          // لا يوجد فرز هنا لأن الترتيب غير مهم للمواعيد كنص
          // allSchedulesWithStatus[currentTeacherName][currentDay].sort((a,b) => getTimeInMinutes(a.timeSlot) - getTimeInMinutes(b.timeSlot));

      } else {
        Logger.log(`Skipping row ${i+1} in Available Schedules due to missing Teacher Name or Day (data: ${JSON.stringify(row)}).`);
      }
    }

    // الآن، نتحقق من المواعيد المحجوزة من شيت "مواعيد الطلبة" ونحدث allSchedulesWithStatus
    const studentSchedulesSheet = spreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
    if (studentSchedulesSheet) {
        const allStudentSchedulesData = studentSchedulesSheet.getDataRange().getValues();
        const bookedSlotsMap = new Map();

        for (let i = STUDENT_SCHEDULE_FIRST_DATA_ROW - 1; i < allStudentSchedulesData.length; i++) {
            const row = allStudentSchedulesData[i];
            const studentId = row[STUDENT_SCHEDULE_STUDENT_ID_COL - 1];
            const studentName = row[STUDENT_SCHEDULE_STUDENT_NAME_COL - 1];
            const schTeacherName = row[STUDENT_SCHEDULE_TEACHER_NAME_COL - 1];

            if (schTeacherName && row[STUDENT_SCHEDULE_DAY1_COL - 1] && row[STUDENT_SCHEDULE_TIME1_COL - 1]) {
                const day = row[STUDENT_SCHEDULE_DAY1_COL - 1].toString().trim();
                const timeText = row[STUDENT_SCHEDULE_TIME1_COL - 1].toString().trim(); // ***** النص كما هو *****
                if (timeText !== '') {
                    const key = `${schTeacherName.toString().trim()}-${day}-${timeText}`;
                    bookedSlotsMap.set(key, { _id: studentId, name: studentName });
                }
            }
            if (schTeacherName && row[STUDENT_SCHEDULE_DAY2_COL - 1] && row[STUDENT_SCHEDULE_TIME2_COL - 1]) {
                const day = row[STUDENT_SCHEDULE_DAY2_COL - 1].toString().trim();
                const timeText = row[STUDENT_SCHEDULE_TIME2_COL - 1].toString().trim(); // ***** النص كما هو *****
                if (timeText !== '') {
                    const key = `${schTeacherName.toString().trim()}-${day}-${timeText}`;
                    bookedSlotsMap.set(key, { _id: studentId, name: studentName });
                }
            }
        }

        for (const teacher in allSchedulesWithStatus) {
            for (const day in allSchedulesWithStatus[teacher]) {
                allSchedulesWithStatus[teacher][day].forEach(slot => {
                    const key = `${slot.teacherName || teacher}-${slot.dayOfWeek || day}-${slot.timeSlot}`;
                    if (bookedSlotsMap.has(key)) {
                        slot.isBooked = true;
                        slot.bookedBy = bookedSlotsMap.get(key);
                    }
                });
            }
        }
    } else {
        Logger.log(`[getTeacherAndAvailableSchedules] Warning: Student schedules sheet '${STUDENT_SCHEDULES_SHEET_NAME}' not found. Cannot determine booked slots.`);
    }

    Logger.log("Available schedules mapped:", JSON.stringify(availableSchedules));
    Logger.log("All schedules with status mapped:", JSON.stringify(allSchedulesWithStatus));
    return JSON.stringify({ teacherNames: teacherNames, availableSchedules: availableSchedules, allSchedulesWithStatus: allSchedulesWithStatus });

  } catch (e) {
    Logger.log(`خطأ عام في getTeacherAndAvailableSchedules: ${e.message}`);
    return JSON.stringify({ teacherNames: [], availableSchedules: {}, allSchedulesWithStatus: {} });
  }
}



/**
 * دالة لإضافة طالب جديد وتسجيل بياناته ومواعيده.
 * @param {Object} studentData كائن يحتوي على بيانات الطالب من الفورم.
 * @returns {Object} حالة العملية (success/error) ومعرف الطالب إذا نجحت.
 */
function addNewStudent(studentData) {
  Logger.log("addNewStudent: Received data:", JSON.stringify(studentData));

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1. إنشاء معرف طالب جديد (Student ID)
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    throw new Error(`شيت المشرف الرئيسي غير موجود: ${MASTER_SHEET_NAME}`);
  }
  const lastRow = masterSheet.getLastRow();
  let newStudentId = "STD001"; // معرف افتراضي لأول طالب

  if (lastRow >= MASTER_FIRST_DATA_ROW) {
    const existingIds = masterSheet.getRange(MASTER_FIRST_DATA_ROW, MASTER_STUDENT_ID_COL, lastRow - MASTER_FIRST_DATA_ROW + 1, 1).getValues()
                            .flat()
                            .filter(id => id && id.toString().startsWith("STD"))
                            .map(id => parseInt(id.toString().replace("STD", ""), 10))
                            .sort((a,b) => a - b);
    if (existingIds.length > 0) {
        newStudentId = "STD" + String(existingIds[existingIds.length - 1] + 1).padStart(3, '0');
    } else {
        // لو في بيانات بس مفيش IDs تبدأ بـ STD (أول ID)
        newStudentId = "STD001";
    }
  } // لو lastRow < MASTER_FIRST_DATA_ROW يبقى مفيش داتا خالص وهنبدأ بـ STD001
  Logger.log(`New Student ID generated: ${newStudentId}`);


  // 2. إضافة بيانات الطالب إلى شيت "بيانات الطلبة"
  const newMasterRow = [
    newStudentId, // معرف الطالب
    studentData.name || '',
    studentData.age || '',
    studentData.number || '',
    '', '', '', // أعمدة فارغة (E, F, G) لو مش مستخدمة في هذا الشيت
    studentData.systemType || '', // النظام (H)
    0, // عدد الحصص الحاضرة (I) - يبدأ بصفر
    'لا يوجد حضور', // حالة التجديد (J) - يبدأ بـ "لا يوجد حضور"
    '', // تاريخ آخر دفع (K)
    0, // المبلغ (L) - يبدأ بصفر
    formatDateOnly(new Date()) // تاريخ الاشتراك (M)
  ];
  masterSheet.appendRow(newMasterRow);
  Logger.log(`Student ${newStudentId} added to Master Sheet.`);


  // 3. إضافة بيانات الطالب إلى شيت "مواعيد الطلبة" (Student Schedules Sheet)
  // وتعديل "المواعيد المتاحة للمعلمين"
  const studentSchedulesSheet = spreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
  if (!studentSchedulesSheet) {
    throw new Error(`شيت مواعيد الطلبة غير موجود: ${STUDENT_SCHEDULES_SHEET_NAME}`);
  }

  // نجيب شيت المواعيد المتاحة مرة واحدة بره اللوب
  const availableSchedulesSheet = spreadsheet.getSheetByName(AVAILABLE_SCHEDULES_SHEET_NAME);
  if (!availableSchedulesSheet) {
    Logger.log(`خطأ: لم يتم العثور على شيت المواعيد المتاحة للمعلمين باسم '${AVAILABLE_SCHEDULES_SHEET_NAME}'. لن يتم مسح المواعيد.`);
    // هنا لا نرمي خطأ عشان بقية العملية تكمل
  }


  const schedules = studentData.schedules;
  const schedulesPerHalfRow = 2; // كل صف هيحتوي على ميعادين
  
  for (let i = 0; i < schedules.length; i += schedulesPerHalfRow) {
    const newScheduleRow = [
      newStudentId,
      studentData.name,
      studentData.age || '',
      studentData.number || '',
      studentData.schedules[0].teacherName, // اسم المعلم الأساسي
    ];

    // الميعاد الأول في الصف
    if (schedules[i]) {
      newScheduleRow[STUDENT_SCHEDULE_DAY1_COL - 1] = schedules[i].day;
      newScheduleRow[STUDENT_SCHEDULE_TIME1_COL - 1] = schedules[i].time;
      
      // ** مسح الميعاد الأول من شيت المواعيد المتاحة **
      // هنستدعي clearScheduledTime لكل ميعاد تم حجزه
      if (availableSchedulesSheet) { 
        clearScheduledTime(availableSchedulesSheet, schedules[i]);
      }
    } else {
      newScheduleRow[STUDENT_SCHEDULE_DAY1_COL - 1] = '';
      newScheduleRow[STUDENT_SCHEDULE_TIME1_COL - 1] = '';
    }

    // الميعاد الثاني في الصف (لو موجود)
    if (schedules[i + 1]) {
      newScheduleRow[STUDENT_SCHEDULE_DAY2_COL - 1] = schedules[i + 1].day;
      newScheduleRow[STUDENT_SCHEDULE_TIME2_COL - 1] = schedules[i + 1].time;

      // ** مسح الميعاد الثاني من شيت المواعيد المتاحة **
      if (availableSchedulesSheet) {
        clearScheduledTime(availableSchedulesSheet, schedules[i + 1]);
      }
    } else {
      newScheduleRow[STUDENT_SCHEDULE_DAY2_COL - 1] = '';
      newScheduleRow[STUDENT_SCHEDULE_TIME2_COL - 1] = '';
    }

    // إضافة البيانات المشتركة
    newScheduleRow[STUDENT_SCHEDULE_SUB_TYPE_COL - 1] = studentData.subscriptionType;
    newScheduleRow[STUDENT_SCHEDULE_SYSTEM_COL - 1] = studentData.systemType;
    newScheduleRow[STUDENT_SCHEDULE_START_DATE_COL - 1] = formatDateOnly(new Date());

    studentSchedulesSheet.appendRow(newScheduleRow);
    Logger.log(`Schedule row added for ${newStudentId} (Batch ${i / schedulesPerHalfRow + 1}): ${JSON.stringify(newScheduleRow)}`);
  }

  // 5. إضافة بيانات الطالب إلى شيت المعلم الخاص بالحضور والغياب (Master Teacher Sheet)
  const uniqueTeacherSheets = new Set(studentData.schedules.map(s => s.teacherName));
  Logger.log(`الطالب ${newStudentId} سيضاف إلى شيتات المعلمين: ${Array.from(uniqueTeacherSheets).join(', ')}`);

  const teachersMainSheet = spreadsheet.getSheetByName(TEACHERS_SHEET_NAME);
  const teachersMainSheetValues = teachersMainSheet.getDataRange().getValues();
  
  for (const teacherName of uniqueTeacherSheets) {
    const teacherRowInMainSheet = teachersMainSheetValues.find(row => row[TEACHER_NAME_COL - 1] === teacherName);
    if (teacherRowInMainSheet) {
      const teacherSheetUrl = teacherRowInMainSheet[TEACHER_SHEET_URL_COL - 1];
      if (teacherSheetUrl) {
        try {
          const teacherSpreadsheet = SpreadsheetApp.openByUrl(teacherSheetUrl);
          const teacherSheet = teacherSpreadsheet.getSheets()[0]; // أول شيت في ملف المعلم

          const newTeacherSheetRow = [
            newStudentId, // معرف الطالب
            studentData.name, // اسم الطالب
            '', '', // عمود فارغ (C), عمود فارغ (D)
            false, false, false, false, false, false, false, false, false, false, false, false // 12 عمود حضور (F-Q)
          ];
          teacherSheet.appendRow(newTeacherSheetRow);
          Logger.log(`Student ${newStudentId} added to teacher sheet: ${teacherSheetUrl}`);
        } catch (e) {
          Logger.log(`خطأ في إضافة الطالب ${newStudentId} إلى شيت المعلم ${teacherName} (${teacherSheetUrl}): ${e.message}`);
        }
      } else {
        Logger.log(`تحذير: لا يوجد رابط شيت للمعلم ${teacherName} في شيت المعلمين الرئيسي.`);
      }
    } else {
      Logger.log(`تحذير: لم يتم العثور على المعلم ${teacherName} في شيت المعلمين الرئيسي.`);
    }
  }

  return { status: "success", studentId: newStudentId };
}


/**
 * دالة مساعدة لمسح ميعاد معين من شيت "المواعيد المتاحة للمعلمين".
 * تم فصلها لتحسين قابلية القراءة والاستخدام المتكرر.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet شيت المواعيد المتاحة.
 * @param {Object} scheduleData كائن يحتوي على (teacherName, day, time) للميعاد المراد مسحه.
 */
function clearScheduledTime(sheet, scheduleData) {
    Logger.log(`[clearScheduledTime] Starting to clear schedule: ${JSON.stringify(scheduleData)}`);

    if (!scheduleData || !scheduleData.teacherName || !scheduleData.day || !scheduleData.time) {
        Logger.log(`[clearScheduledTime] ERROR: Invalid scheduleData received: ${JSON.stringify(scheduleData)}`);
        return;
    }

    const availableValues = sheet.getDataRange().getDisplayValues();
    let foundAndCleared = false;

    Logger.log(`[clearScheduledTime] Sheet read. Rows: ${availableValues.length}, Cols: ${availableValues[0] ? availableValues[0].length : 0}`);
    Logger.log(`[clearScheduledTime] Looking for Teacher: '${scheduleData.teacherName}', Day: '${scheduleData.day}', Time: '${scheduleData.time}'`);


    for (let r = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; r < availableValues.length; r++) {
      const currentRow = availableValues[r];
      if (currentRow.length < AVAILABLE_SCHEDULE_TIMES_START_COL) {
          Logger.log(`[clearScheduledTime] Skipping row ${r+1}: Too few columns. Row content: ${JSON.stringify(currentRow)}`);
          continue;
      }

      const currentTeacherNameInSheet = currentRow[AVAILABLE_SCHEDULE_TEACHER_NAME_COL - 1];
      const currentDayInSheet = currentRow[AVAILABLE_SCHEDULE_DAY_COL - 1];

      Logger.log(`[clearScheduledTime] Checking row ${r+1}: Sheet Teacher: '${currentTeacherNameInSheet}', Sheet Day: '${currentDayInSheet}'`);

      if (currentTeacherNameInSheet && currentTeacherNameInSheet.toString().trim() === scheduleData.teacherName.toString().trim() &&
          currentDayInSheet && currentDayInSheet.toString().trim() === scheduleData.day.toString().trim()) {

        Logger.log(`[clearScheduledTime] MATCH found for Teacher/Day in row ${r+1}. Now checking time slots.`);

        // نمر على أعمدة المواعيد في الصف ده عشان نلاقي الميعاد المحدد
        for (let col = AVAILABLE_SCHEDULE_TIMES_START_COL - 1; col < currentRow.length; col++) {
          const timeSlotInCell = currentRow[col];

          Logger.log(`[clearScheduledTime] Row ${r+1}, Col ${col+1} (Value: '${timeSlotInCell}'). Comparing with target time: '${scheduleData.time}'`);

          if (timeSlotInCell && timeSlotInCell.toString().trim() === scheduleData.time.toString().trim()) {
            Logger.log(`[clearScheduledTime] EXACT MATCH FOUND! Clearing cell (Row: ${r+1}, Col: ${col+1}).`);
            sheet.getRange(r + 1, col + 1).clearContent();
            foundAndCleared = true;
            break;
          }
        }
      }
      if (foundAndCleared) {
        break;
      }
    }

    if (!foundAndCleared) {
      Logger.log(`[clearScheduledTime] WARNING: Schedule not found for clearing: Teacher '${scheduleData.teacherName}', Day '${scheduleData.day}', Time '${scheduleData.time}'.`);
    }
}


/**
 * للبحث عن طالب في شيت بيانات الطلبة بالمعرف أو رقم الهاتف.
 * @param {string} query المعرف أو رقم الهاتف للبحث عنه.
 * @returns {string} JSON string containing student data if found, otherwise null.
 */
function searchStudent(query) {
  Logger.log(`[searchStudent] Searching for: '${query}'`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    Logger.log(`[searchStudent] Error: Master sheet '${MASTER_SHEET_NAME}' not found.`);
    return null;
  }

  const values = masterSheet.getDataRange().getValues();
  // بما أننا بنبحث بجزء، ممكن نلاقي أكتر من طالب، عشان كده هنعمل مصفوفة للنتائج
  const foundStudents = []; 

  // تحضير الـ query للبحث المرن
  const lowerCaseQuery = query.toLowerCase().trim();
  let numericQuery = null;
  if (!isNaN(parseFloat(lowerCaseQuery)) && isFinite(lowerCaseQuery)) {
      numericQuery = lowerCaseQuery; // لو البحث رقمي (زي رقم تليفون أو جزء من ID)
  }
  const queryWithoutStd = lowerCaseQuery.startsWith("std") ? lowerCaseQuery.replace("std", "") : lowerCaseQuery;


  for (let i = MASTER_FIRST_DATA_ROW - 1; i < values.length; i++) {
    const row = values[i];
    const studentId = row[MASTER_STUDENT_ID_COL - 1] ? row[MASTER_STUDENT_ID_COL - 1].toString().trim() : '';
    const studentNumber = row[MASTER_STUDENT_NUMBER_COL - 1] ? row[MASTER_STUDENT_NUMBER_COL - 1].toString().trim() : '';
    const studentName = row[MASTER_STUDENT_NAME_COL - 1] ? row[MASTER_STUDENT_NAME_COL - 1].toString().trim() : '';
    const studentSystem = row[MASTER_SYSTEM_COL - 1] ? row[MASTER_SYSTEM_COL - 1].toString().trim() : '';

    // معالجة معرف الطالب للبحث بجزء (STD001 -> 001)
    const studentIdWithoutStd = studentId.startsWith("STD") ? studentId.toLowerCase().replace("std", "") : studentId.toLowerCase();

    // شروط البحث المرنة:
    const isIdMatch = studentId.toLowerCase().includes(lowerCaseQuery) || // بحث بجزء من ID كاملاً
                      studentIdWithoutStd.includes(queryWithoutStd);     // بحث بجزء من الرقم فقط (بعد إزالة STD)
    const isNumberMatch = studentNumber.includes(lowerCaseQuery);      // بحث بجزء من رقم الهاتف
    const isNameMatch = studentName.toLowerCase().includes(lowerCaseQuery); // (لو عايزين نبحث بالاسم كمان)
    const isSystemMatch = studentSystem.toLowerCase().includes(lowerCaseQuery); // (لو عايزين نبحث بالنظام كمان)


    if (isIdMatch || isNumberMatch || isNameMatch || isSystemMatch) { // لو عايزين ندمج البحث بالاسم والنظام كمان
      // جلب بيانات الطالب الأساسية
      const studentInfo = {
        id: studentId,
        name: studentName,
        age: row[MASTER_STUDENT_AGE_COL - 1] || '',
        number: studentNumber,
        system: studentSystem,
        subscriptionType: row[MASTER_SUB_TYPE_COL - 1] || '', // تأكد من أن هذا الثابت موجود لو هتستخدمه
        attendance: row[MASTER_TOTAL_ATTENDANCE_COL - 1] || 0,
        renewalStatus: row[MASTER_RENEWAL_STATUS_COL - 1] || '',
        lastPaymentDate: row[MASTER_LAST_PAYMENT_DATE_COL - 1] ? formatDateOnly(new Date(row[MASTER_LAST_PAYMENT_DATE_COL - 1])) : '',
        amount: row[MASTER_AMOUNT_COL - 1] || 0,
        subscriptionDate: row[MASTER_SUBSCRIPTION_DATE_COL - 1] ? formatDateOnly(new Date(row[MASTER_SUBSCRIPTION_DATE_COL - 1])) : '',
        masterRow: i + 1 // رقم الصف الفعلي في الشيت الرئيسي
      };
      
      // جلب مواعيد الطالب من شيت مواعيد الطلبة
      const studentSchedules = getStudentSchedulesInternal(studentId); // دالة داخلية نستخدمها هنا
      studentInfo.schedules = studentSchedules;
      
      foundStudents.push(studentInfo); // نضيف الطالب اللي لقيناه لمصفوفة النتائج
    }
  }

  Logger.log(`[searchStudent] Found ${foundStudents.length} students matching query '${query}'.`);
  // هنرجع أول طالب نلاقيه (أو الأنسب لو فيه منطق ترتيب)
  // لو عايزين نرجع كل النتائج للواجهة عشان المستخدم يختار، هنرجع foundStudents كلها
  // لكن للتبسيط ولأن الواجهة مصممة لعرض طالب واحد، هنرجع أول طالب أو null لو مفيش
  return foundStudents.length > 0 ? JSON.stringify(foundStudents[0]) : null;
}



/**
 * دالة داخلية لجلب جميع المواعيد الخاصة بطالب معين من شيت مواعيد الطلبة.
 * تُستخدم داخلياً بواسطة دوال أخرى في Apps Script.
 * @param {string} studentId معرف الطالب.
 * @returns {Array<Object>} مصفوفة من كائنات المواعيد للطالب.
 */
function getStudentSchedulesInternal(studentId) {
  Logger.log(`[getStudentSchedulesInternal] Fetching schedules for student ID: ${studentId}`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const studentSchedulesSheet = spreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
  if (!studentSchedulesSheet) {
    Logger.log(`[getStudentSchedulesInternal] Error: Student schedules sheet '${STUDENT_SCHEDULES_SHEET_NAME}' not found.`);
    return [];
  }

  const allSchedulesData = studentSchedulesSheet.getDataRange().getValues();
  const studentSchedules = [];

  for (let i = STUDENT_SCHEDULE_FIRST_DATA_ROW - 1; i < allSchedulesData.length; i++) {
    const row = allSchedulesData[i];
    const currentStudentId = row[STUDENT_SCHEDULE_STUDENT_ID_COL - 1];

    if (currentStudentId && currentStudentId.toString().trim() === studentId.trim()) {
      // جمع الميعاد الأول
      if (row[STUDENT_SCHEDULE_DAY1_COL - 1] && row[STUDENT_SCHEDULE_TIME1_COL - 1]) {
        studentSchedules.push({
          teacherName: row[STUDENT_SCHEDULE_TEACHER_NAME_COL - 1] || '',
          day: row[STUDENT_SCHEDULE_DAY1_COL - 1].toString().trim(),
          time: row[STUDENT_SCHEDULE_TIME1_COL - 1].toString().trim(), // ***** النص كما هو *****
          scheduleRow: i + 1 // رقم الصف في شيت مواعيد الطلبة
        });
      }
      // جمع الميعاد الثاني
      if (row[STUDENT_SCHEDULE_DAY2_COL - 1] && row[STUDENT_SCHEDULE_TIME2_COL - 1]) {
        studentSchedules.push({
          teacherName: row[STUDENT_SCHEDULE_TEACHER_NAME_COL - 1] || '', // نفس المعلم لليومين
          day: row[STUDENT_SCHEDULE_DAY2_COL - 1].toString().trim(),
          time: row[STUDENT_SCHEDULE_TIME2_COL - 1].toString().trim(), // ***** النص كما هو *****
          scheduleRow: i + 1 // رقم الصف في شيت مواعيد الطلبة
        });
      }
    }
  }
  Logger.log(`[getStudentSchedulesInternal] Found ${studentSchedules.length} schedules for student ${studentId}.`);
  return studentSchedules;
}





/**
 * دالة مساعدة لإعادة ميعاد محجوز إلى شيت "المواعيد المتاحة للمعلمين".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet شيت المواعيد المتاحة.
 * @param {Object} scheduleData كائن يحتوي على (teacherName, day, time) للميعاد المراد إعادته.
 */
function returnScheduledTime(sheet, scheduleData) {
  Logger.log(`[returnScheduledTime] Attempting to return schedule: ${JSON.stringify(scheduleData)}`);

  if (!scheduleData || !scheduleData.teacherName || !scheduleData.day || !scheduleData.time) {
    Logger.log(`[returnScheduledTime] ERROR: Invalid scheduleData received: ${JSON.stringify(scheduleData)}`);
    return;
  }

  const availableValues = sheet.getDataRange().getDisplayValues();
  let foundAndReturned = false;

  Logger.log(`[returnScheduledTime] Sheet read. Rows: ${availableValues.length}, Cols: ${availableValues[0] ? availableValues[0].length : 0}`);
  Logger.log(`[returnScheduledTime] Looking for slot for Teacher: '${scheduleData.teacherName}', Day: '${scheduleData.day}', Time: '${scheduleData.time}'`);

  for (let r = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; r < availableValues.length; r++) {
    const currentRow = availableValues[r];
    if (currentRow.length < AVAILABLE_SCHEDULE_TIMES_START_COL) {
        continue;
    }

    const currentTeacherNameInSheet = currentRow[AVAILABLE_SCHEDULE_TEACHER_NAME_COL - 1];
    const currentDayInSheet = currentRow[AVAILABLE_SCHEDULE_DAY_COL - 1];

    if (currentTeacherNameInSheet && currentTeacherNameInSheet.toString().trim() === scheduleData.teacherName.toString().trim() &&
        currentDayInSheet && currentDayInSheet.toString().trim() === scheduleData.day.toString().trim()) {

      Logger.log(`[returnScheduledTime] MATCH found for Teacher/Day in row ${r+1}. Now checking time slots to return.`);

      // هنا بقى نمر على أعمدة المواعيد ونبحث عن أول خلية فارغة نكتب فيها الميعاد
      for (let col = AVAILABLE_SCHEDULE_TIMES_START_COL - 1; col < currentRow.length; col++) {
        const timeSlotInCell = currentRow[col];

        if (!timeSlotInCell || timeSlotInCell.toString().trim() === "") { // لو الخلية فاضية
            sheet.getRange(r + 1, col + 1).setValue(scheduleData.time);
            Logger.log(`[returnScheduledTime] Returned schedule '${scheduleData.time}' to empty cell (Row: ${r+1}, Col: ${col+1}).`);
            foundAndReturned = true;
            break;
        } else if (timeSlotInCell.toString().trim() === scheduleData.time.toString().trim()) {
            Logger.log(`[returnScheduledTime] WARNING: Schedule '${scheduleData.time}' already exists in cell (Row: ${r+1}, Col: ${col+1}). Not returning.`);
            foundAndReturned = true;
            break;
        }
      }
    }
    if (foundAndReturned) {
      break;
    }
  }

  if (!foundAndReturned) {
    Logger.log(`[returnScheduledTime] WARNING: Schedule not returned: Teacher '${scheduleData.teacherName}', Day '${scheduleData.day}', Time '${scheduleData.time}'. Could not find an empty slot.`);
  }
}



/**
 * دالة لتعديل بيانات طالب موجود ومواعيده.
 * @param {Object} studentData كائن يحتوي على بيانات الطالب المعدلة من الفورم.
 * يجب أن يحتوي على studentData.id و studentData.masterRow
 * و studentData.schedules (المواعيد الجديدة).
 * @returns {Object} حالة العملية (success/error).
 */
function updateStudent(studentData) {
  Logger.log(`[updateStudent] Updating student: ${studentData.id}, Row: ${studentData.masterRow}`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 1. تحديث بيانات الطالب في شيت "بيانات الطلبة"
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    throw new Error(`شيت المشرف الرئيسي غير موجود: ${MASTER_SHEET_NAME}`);
  }
  const masterRowRange = masterSheet.getRange(studentData.masterRow, 1, 1, masterSheet.getLastColumn());
  const currentMasterRow = masterRowRange.getValues()[0];

  currentMasterRow[MASTER_STUDENT_NAME_COL - 1] = studentData.name || '';
  currentMasterRow[MASTER_STUDENT_AGE_COL - 1] = studentData.age || '';
  currentMasterRow[MASTER_STUDENT_NUMBER_COL - 1] = studentData.number || '';
  currentMasterRow[MASTER_SYSTEM_COL - 1] = studentData.systemType || '';
  // تأكد من أن هذه الأعمدة موجودة في studentData لو كانت قابلة للتعديل من الواجهة
  currentMasterRow[MASTER_SUBSCRIPTION_DATE_COL - 1] = studentData.subscriptionDate || currentMasterRow[MASTER_SUBSCRIPTION_DATE_COL - 1];
  // أعمدة الحضور والتجديد والمبلغ وتاريخ آخر دفع لا يتم تعديلها هنا

  masterRowRange.setValues([currentMasterRow]);
  Logger.log(`[updateStudent] Master sheet updated for student ${studentData.id}.`);

  // 2. معالجة مواعيد الطالب (إرجاع القديمة، حذف القديمة، إضافة الجديدة، خصم الجديدة)
  const studentSchedulesSheet = spreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
  if (!studentSchedulesSheet) {
    throw new Error(`شيت مواعيد الطلبة غير موجود: ${STUDENT_SCHEDULES_SHEET_NAME}`);
  }
  const availableSchedulesSheet = spreadsheet.getSheetByName(AVAILABLE_SCHEDULES_SHEET_NAME);
  if (!availableSchedulesSheet) {
    Logger.log(`[updateStudent] Warning: Available Schedules sheet '${AVAILABLE_SCHEDULES_SHEET_NAME}' not found. Cannot return/clear schedules.`);
  }

  // 2.1. جلب المواعيد القديمة للطالب من شيت "مواعيد الطلبة"
  const oldSchedules = getStudentSchedulesInternal(studentData.id);
  Logger.log(`[updateStudent] Old schedules found for ${studentData.id}: ${JSON.stringify(oldSchedules)}`);

  // 2.2. تحديد المواعيد القديمة التي لم تعد موجودة في المواعيد الجديدة لإعادتها
  const newSchedulesSet = new Set(studentData.schedules.map(s => `<span class="math-inline">\{s\.teacherName\}\-</span>{s.day}-${s.time}`));
  Logger.log(`[updateStudent] New schedules (unique set): ${JSON.stringify(Array.from(newSchedulesSet))}`);

  if (availableSchedulesSheet) {
      oldSchedules.forEach(oldSchedule => {
          const oldScheduleString = `<span class="math-inline">\{oldSchedule\.teacherName\}\-</span>{oldSchedule.day}-${oldSchedule.time}`;
          if (!newSchedulesSet.has(oldScheduleString)) {
              // الميعاد ده كان موجود في القديم بس مش موجود في الجديد، يبقى نرجعه
              returnScheduledTime(availableSchedulesSheet, oldSchedule);
          }
      });
  }

  // 2.3. حذف جميع المواعيد القديمة للطالب من شيت "مواعيد الطلبة"
  // لازم نمسح الصفوف من تحت لفوق عشان الفهارس ما تتلخبطش
  const allStudentSchedulesData = studentSchedulesSheet.getDataRange().getValues();
  let rowsToDelete = [];
  for (let i = allStudentSchedulesData.length - 1; i >= STUDENT_SCHEDULE_FIRST_DATA_ROW - 1; i--) {
      if (allStudentSchedulesData[i][STUDENT_SCHEDULE_STUDENT_ID_COL - 1] === studentData.id) {
          rowsToDelete.push(i + 1); // 1-indexed row number
      }
  }
  rowsToDelete.forEach(rowNum => {
      studentSchedulesSheet.deleteRow(rowNum);
      Logger.log(`[updateStudent] Deleted old schedule row for ${studentData.id} at row ${rowNum} from Student Schedules Sheet.`);
  });
  

  // 2.4. إضافة المواعيد الجديدة للطالب في شيت "مواعيد الطلبة" (بنفس منطق addNewStudent)
  // وخصمها من شيت "المواعيد المتاحة للمعلمين"
  const schedulesToAddNew = studentData.schedules;
  const schedulesPerHalfRow = 2; 

  for (let i = 0; i < schedulesToAddNew.length; i += schedulesPerHalfRow) {
    const newScheduleRowData = [
      studentData.id,
      studentData.name,
      studentData.age || '',
      studentData.number || '',
      schedulesToAddNew[0].teacherName, // اسم المعلم الأساسي
    ];

    // الميعاد الأول في الصف
    if (schedulesToAddNew[i]) {
      newScheduleRowData[STUDENT_SCHEDULE_DAY1_COL - 1] = schedulesToAddNew[i].day;
      newScheduleRowData[STUDENT_SCHEDULE_TIME1_COL - 1] = schedulesToAddNew[i].time;
      if (availableSchedulesSheet) { 
        clearScheduledTime(availableSchedulesSheet, schedulesToAddNew[i]);
      }
    } else {
      newScheduleRowData[STUDENT_SCHEDULE_DAY1_COL - 1] = '';
      newScheduleRowData[STUDENT_SCHEDULE_TIME1_COL - 1] = '';
    }

    // الميعاد الثاني في الصف (لو موجود)
    if (schedulesToAddNew[i + 1]) {
      newScheduleRowData[STUDENT_SCHEDULE_DAY2_COL - 1] = schedulesToAddNew[i + 1].day;
      newScheduleRowData[STUDENT_SCHEDULE_TIME2_COL - 1] = schedulesToAddNew[i + 1].time;
      if (availableSchedulesSheet) {
        clearScheduledTime(availableSchedulesSheet, schedulesToAddNew[i + 1]);
      }
    } else {
      newScheduleRowData[STUDENT_SCHEDULE_DAY2_COL - 1] = '';
      newScheduleRowData[STUDENT_SCHEDULE_TIME2_COL - 1] = '';
    }

    newScheduleRowData[STUDENT_SCHEDULE_SUB_TYPE_COL - 1] = studentData.subscriptionType;
    newScheduleRowData[STUDENT_SCHEDULE_SYSTEM_COL - 1] = studentData.systemType;
    newScheduleRowData[STUDENT_SCHEDULE_START_DATE_COL - 1] = formatDateOnly(new Date());

    studentSchedulesSheet.appendRow(newScheduleRowData);
    Logger.log(`[updateStudent] Added new schedule row for ${studentData.id}: ${JSON.stringify(newScheduleRowData)}`);
  }

  return { status: "success" };
}




/**
 * دالة لحذف طالب من النظام (نقله للأرشيف وإعادة مواعيده).
 * @param {string} studentId معرف الطالب المراد حذفه.
 * @param {string} reason سبب الحذف (اختياري).
 * @returns {Object} حالة العملية (success/error).
 */
function deleteStudent(studentId, reason = "") {
  Logger.log(`[deleteStudent] Attempting to delete student: ${studentId}, Reason: ${reason}`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  const studentSchedulesSheet = spreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
  const archiveSheet = spreadsheet.getSheetByName(ARCHIVE_SHEET_NAME);
  const availableSchedulesSheet = spreadsheet.getSheetByName(AVAILABLE_SCHEDULES_SHEET_NAME);

  if (!masterSheet) throw new Error(`Master sheet '${MASTER_SHEET_NAME}' not found.`);
  if (!studentSchedulesSheet) throw new Error(`Student schedules sheet '${STUDENT_SCHEDULES_SHEET_NAME}' not found.`);
  if (!archiveSheet) throw new Error(`Archive sheet '${ARCHIVE_SHEET_NAME}' not found.`);
  if (!availableSchedulesSheet) Logger.log(`[deleteStudent] Warning: Available Schedules sheet '${AVAILABLE_SCHEDULES_SHEET_NAME}' not found. Cannot return schedules.`);

  // 1. جلب المواعيد الحالية للطالب لإعادتها
  const studentSchedules = getStudentSchedulesInternal(studentId);
  Logger.log(`[deleteStudent] Found ${studentSchedules.length} schedules for student ${studentId} to return.`);

  if (availableSchedulesSheet) {
      studentSchedules.forEach(schedule => {
          returnScheduledTime(availableSchedulesSheet, schedule);
      });
  }
  Logger.log(`[deleteStudent] Attempted to return all schedules for student ${studentId}.`);

  // 2. نقل بيانات الطالب من مواعيد الطلبة إلى الأرشيف
  // بما أن الأرشيف بيحتوي على نفس أعمدة مواعيد الطلبة + 2 (تاريخ وسبب الحذف)
  const allStudentSchedulesData = studentSchedulesSheet.getDataRange().getValues();
  const rowsToDeleteFromStudentSchedules = [];

  for (let i = allStudentSchedulesData.length - 1; i >= STUDENT_SCHEDULE_FIRST_DATA_ROW - 1; i--) {
      const row = allStudentSchedulesData[i];
      if (row[STUDENT_SCHEDULE_STUDENT_ID_COL - 1] === studentId) {
          const rowToArchive = [...row]; // نسخ الصف كما هو

          // إضافة تاريخ وسبب الحذف (مرة واحدة فقط في أول صف للطالب)
          if (rowsToDeleteFromStudentSchedules.length === 0) { // لو ده أول صف بنضيفه للطالب
              rowToArchive[ARCHIVE_DATE_DELETED_COL - 1] = formatDateOnly(new Date());
              rowToArchive[ARCHIVE_REASON_DELETED_COL - 1] = reason;
          } else {
              rowToArchive[ARCHIVE_DATE_DELETED_COL - 1] = ''; // باقي الصفوف هتكون فاضية
              rowToArchive[ARCHIVE_REASON_DELETED_COL - 1] = '';
          }
          archiveSheet.appendRow(rowToArchive);
          Logger.log(`[deleteStudent] Archived schedule row for ${studentId} from Student Schedules Sheet.`);
          rowsToDeleteFromStudentSchedules.push(i + 1); // 1-indexed row number
      }
  }
  
  // 3. حذف الطالب من شيت "بيانات الطلبة" الرئيسي
  const masterValues = masterSheet.getDataRange().getValues();
  let masterRowToDelete = -1;
  for (let i = MASTER_FIRST_DATA_ROW - 1; i < masterValues.length; i++) {
      if (masterValues[i][MASTER_STUDENT_ID_COL - 1] === studentId) {
          masterRowToDelete = i + 1;
          break;
      }
  }
  if (masterRowToDelete !== -1) {
      masterSheet.deleteRow(masterRowToDelete);
      Logger.log(`[deleteStudent] Student ${studentId} deleted from Master Sheet.`);
  } else {
      Logger.log(`[deleteStudent] Warning: Student ${studentId} not found in Master Sheet for deletion.`);
  }

  // 4. حذف صفوف الطالب من شيت "مواعيد الطلبة"
  // بما أننا جمعنا rowsToDeleteFromStudentSchedules من تحت لفوق، هنمسح بنفس الترتيب
  rowsToDeleteFromStudentSchedules.sort((a,b) => b - a).forEach(rowNum => { // لازم نرتب تنازلي عشان المسح مايلخبطش الفهارس
      studentSchedulesSheet.deleteRow(rowNum);
      Logger.log(`[deleteStudent] Deleted schedule row for ${studentId} at row ${rowNum} from Student Schedules Sheet.`);
  });

  return { status: "success" };
}



/**
 * تجلب جميع المواعيد الخاصة بمعلم معين من شيت "المواعيد المتاحة للمعلمين"
 * وتحدد ما إذا كانت هذه المواعيد محجوزة من قبل طلاب.
 * تُستخدم لملء المودال في صفحة تحديث مواعيد المعلمين.
 * @param {string} teacherName اسم المعلم المراد جلب مواعيده.
 * @returns {string} JSON string containing an array of schedule slots with their status.
 * مثال: [{ dayOfWeek: "الاحد", timeSlot: "09:00 ص", isBooked: true, bookedBy: { _id: "STD001", name: "أحمد" } }, ...]
 */
function getTeacherSchedulesForUpdate(teacherName) {
  Logger.log(`[getTeacherSchedulesForUpdate] Fetching schedules for teacher: '${teacherName}'`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  const availableSchedulesSheet = spreadsheet.getSheetByName(AVAILABLE_SCHEDULES_SHEET_NAME);
  if (!availableSchedulesSheet) {
    Logger.log(`[getTeacherSchedulesForUpdate] Error: Available Schedules sheet '${AVAILABLE_SCHEDULES_SHEET_NAME}' not found.`);
    return JSON.stringify([]);
  }

  const studentSchedulesSheet = spreadsheet.getSheetByName(STUDENT_SCHEDULES_SHEET_NAME);
  if (!studentSchedulesSheet) {
      Logger.log(`[getTeacherSchedulesForUpdate] Warning: Student Schedules sheet '${STUDENT_SCHEDULES_SHEET_NAME}' not found. Cannot determine booked slots.`);
  }

  const allAvailableTimeSlots = []; // لتخزين كل المواعيد الممكنة للمعلم مع حالتها

  const allScheduleData = availableSchedulesSheet.getDataRange().getDisplayValues(); 

  if (allScheduleData.length < AVAILABLE_SCHEDULE_FIRST_DATA_ROW) {
      Logger.log("[getTeacherSchedulesForUpdate] Available Schedules sheet has no data rows.");
      return JSON.stringify([]);
  }

  // جلب كل المواعيد المحتملة من جميع صفوف البيانات (لاستخدامها كـ "رؤوس" لكل يوم)
  let allPossibleTimeSlots = new Set();
  for (let i = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; i < allScheduleData.length; i++) {
      const row = allScheduleData[i];
      for (let col = AVAILABLE_SCHEDULE_TIMES_START_COL - 1; col < row.length; col++) {
          const timeValue = row[col];
          if (timeValue && timeValue.toString().trim() !== "") {
              allPossibleTimeSlots.add(timeValue.toString().trim());
          }
      }
  }
  const sortedPossibleTimeSlots = Array.from(allPossibleTimeSlots).sort(); // ترتيب المواعيد


  // نمر على بيانات المعلمين المتاحة من الصف اللي فيه أول معلم
  for (let i = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; i < allScheduleData.length; i++) {
    const row = allScheduleData[i];
    const currentTeacherNameInSheet = row[AVAILABLE_SCHEDULE_TEACHER_NAME_COL - 1]; 
    const currentDayInSheet = row[AVAILABLE_SCHEDULE_DAY_COL - 1]; 
    
    // لو ده صف المعلم اللي بنبحث عنه
    if (currentTeacherNameInSheet && currentTeacherNameInSheet.toString().trim() === teacherName.trim() &&
        currentDayInSheet && currentDayInSheet.toString().trim() !== "") {
      
      const currentDay = currentDayInSheet.toString().trim();

      // نمر على كل المواعيد الممكنة اللي جمعناها
      for (let k = 0; k < sortedPossibleTimeSlots.length; k++) {
          const expectedTime = sortedPossibleTimeSlots[k]; // الميعاد المتوقع (من قائمة كل المواعيد)
          let isTimeSlotAvailableInThisRow = false;
          let colIndex = -1;

          // نبحث عن الميعاد ده في الصف الحالي للمعلم واليوم
          for (let j = AVAILABLE_SCHEDULE_TIMES_START_COL - 1; j < row.length; j++) {
              if (row[j] && row[j].toString().trim() === expectedTime) {
                  isTimeSlotAvailableInThisRow = true;
                  colIndex = j;
                  break;
              }
          }

          if (isTimeSlotAvailableInThisRow) {
              // الميعاد ده متاح حاليا في شيت المواعيد المتاحة
              allAvailableTimeSlots.push({
                  teacherName: teacherName,
                  dayOfWeek: currentDay,
                  timeSlot: expectedTime,
                  isBooked: false,
                  bookedBy: null
              });
          }
      }
    }
  }

  // الآن، نتحقق من المواعيد المحجوزة من شيت "مواعيد الطلبة"
  if (studentSchedulesSheet) {
      const allStudentSchedulesData = studentSchedulesSheet.getDataRange().getValues();
      const bookedSlotsMap = new Map(); // Map<"Teacher-Day-Time", {studentId, studentName}>

      for (let i = STUDENT_SCHEDULE_FIRST_DATA_ROW - 1; i < allStudentSchedulesData.length; i++) {
          const row = allStudentSchedulesData[i];
          const studentId = row[STUDENT_SCHEDULE_STUDENT_ID_COL - 1];
          const studentName = row[STUDENT_SCHEDULE_STUDENT_NAME_COL - 1];
          const schTeacherName = row[STUDENT_SCHEDULE_TEACHER_NAME_COL - 1];

          // الميعاد الأول في الصف
          if (row[STUDENT_SCHEDULE_DAY1_COL - 1] && row[STUDENT_SCHEDULE_TIME1_COL - 1]) {
              const day = row[STUDENT_SCHEDULE_DAY1_COL - 1].toString().trim();
              const time = row[STUDENT_SCHEDULE_TIME1_COL - 1].toString().trim();
              const key = `${schTeacherName}-${day}-${time}`;
              bookedSlotsMap.set(key, { _id: studentId, name: studentName });
          }
          // الميعاد الثاني في الصف
          if (row[STUDENT_SCHEDULE_DAY2_COL - 1] && row[STUDENT_SCHEDULE_TIME2_COL - 1]) {
              const day = row[STUDENT_SCHEDULE_DAY2_COL - 1].toString().trim();
              const time = row[STUDENT_SCHEDULE_TIME2_COL - 1].toString().trim();
              const key = `${schTeacherName}-${day}-${time}`;
              bookedSlotsMap.set(key, { _id: studentId, name: studentName });
          }
      }

      // نحدث حالة isBooked بناءً على المواعيد المحجوزة
      allAvailableTimeSlots.forEach(slot => {
          const key = `${slot.teacherName}-${slot.dayOfWeek}-${slot.timeSlot}`;
          if (bookedSlotsMap.has(key)) {
              slot.isBooked = true;
              slot.bookedBy = bookedSlotsMap.get(key);
          }
      });
  }
  
  Logger.log(`[getTeacherSchedulesForUpdate] Returning ${allAvailableTimeSlots.length} slots for '${teacherName}'.`);
  return JSON.stringify(allAvailableTimeSlots);
}



/**
 * دالة لتحديث المواعيد المتاحة لمعلم معين في يوم معين.
 * ستقوم بملء المواعيد الجديدة في الأعمدة الفارغة، وتخطي الأعمدة المملوءة،
 * وإضافة أعمدة جديدة إذا لزم الأمر، دون اشتراط الترتيب الزمني للمواعيد المكتوبة.
 *
 * @param {string} teacherName اسم المعلم.
 * @param {string} day اليوم (مثلاً: "السبت").
 * @param {Array<string>} newTimeSlots مصفوفة بالمواعيد النصية الجديدة لهذا اليوم (مثلاً: ["09:00 ص", "10:00 ص"]).
 * @returns {Object} حالة العملية (success/error).
 */
function updateTeacherAvailableSchedules(teacherName, day, newTimeSlots) {
  Logger.log(`[updateTeacherAvailableSchedules] Starting: Teacher: '${teacherName}', Day: '${day}'. New slots: ${JSON.stringify(newTimeSlots)}`);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const availableSchedulesSheet = spreadsheet.getSheetByName(AVAILABLE_SCHEDULES_SHEET_NAME);
  if (!availableSchedulesSheet) {
    Logger.log(`[updateTeacherAvailableSchedules] Error: Available Schedules sheet '${AVAILABLE_SCHEDULES_SHEET_NAME}' not found.`);
    return { status: "error", message: `شيت المواعيد المتاحة للمعلمين غير موجود: ${AVAILABLE_SCHEDULES_SHEET_NAME}` };
  }

  // نحصل على البيانات الحالية للصف الذي سنعدله فقط (لكي نقلل القراءات)
  const allSheetData = availableSchedulesSheet.getDataRange().getDisplayValues(); // لنجد الصف أولاً
  let teacherDayRowIndex = -1;

  // 1. نلاقي صف المعلم واليوم المحدد
  for (let i = AVAILABLE_SCHEDULE_FIRST_DATA_ROW - 1; i < allSheetData.length; i++) {
    const row = allSheetData[i]; // نستخدم allSheetData هنا
    const currentTeacherNameInSheet = row[AVAILABLE_SCHEDULE_TEACHER_NAME_COL - 1];
    const currentDayInSheet = row[AVAILABLE_SCHEDULE_DAY_COL - 1];

    if (currentTeacherNameInSheet && currentTeacherNameInSheet.toString().trim() === teacherName.trim() &&
        currentDayInSheet && currentDayInSheet.toString().trim() === day.trim()) {
      teacherDayRowIndex = i;
      break;
    }
  }

  if (teacherDayRowIndex === -1) {
    Logger.log(`[updateTeacherAvailableSchedules] Error: Teacher '${teacherName}' or Day '${day}' not found in Available Schedules sheet.`);
    return { status: "error", message: `لم يتم العثور على المعلم أو اليوم لتحديث المواعيد.` };
  }

  const rowToUpdate = teacherDayRowIndex + 1; // 1-indexed row number
  // نقرأ الصف المحدد لنتعامل معه
  const currentRowValues = availableSchedulesSheet.getRange(rowToUpdate, 1, 1, availableSchedulesSheet.getMaxColumns()).getValues()[0];
  const currentRowDisplayValues = availableSchedulesSheet.getRange(rowToUpdate, 1, 1, availableSchedulesSheet.getMaxColumns()).getDisplayValues()[0];


  const startColumnForTimes = AVAILABLE_SCHEDULE_TIMES_START_COL - 1; // 0-indexed column index for time slots

  // Set (مجموعة) للمواعيد الجديدة لسهولة البحث فيها
  const newTimeSlotsSet = new Set(newTimeSlots.map(slot => slot.toString().trim()));

  // مصفوفة لتخزين الخلايا التي يجب مسحها (المواعيد القديمة التي لم تعد موجودة في القائمة الجديدة)
  const rangesToClear = [];

  // مصفوفة لتخزين المواعيد الجديدة التي يجب إضافتها
  const slotsToAdd = [];

  // مجموعة لتتبع المواعيد الموجودة حالياً في الصف والتي هي أيضاً في القائمة الجديدة (لا تمسح ولا تضاف)
  const timesAlreadyInRowAndNewList = new Set();


  // 1. نمر على المواعيد الموجودة حالياً في الصف ونحدد ما يجب مسحه وما يجب إضافته
  for (let col = startColumnForTimes; col < currentRowDisplayValues.length; col++) {
    const timeInCell = currentRowDisplayValues[col]; // القيمة المعروضة في الخلية

    if (timeInCell && timeInCell.toString().trim() !== "") { // إذا كانت الخلية ليست فارغة
      const trimmedTimeInCell = timeInCell.toString().trim();
      if (newTimeSlotsSet.has(trimmedTimeInCell)) {
        // الميعاد ده موجود في الصف القديم وموجود في القائمة الجديدة
        timesAlreadyInRowAndNewList.add(trimmedTimeInCell);
      } else {
        // الميعاد ده موجود في الصف القديم لكن مش في القائمة الجديدة، يبقى لازم نمسحه
        rangesToClear.push(availableSchedulesSheet.getRange(rowToUpdate, col + 1));
      }
    }
  }

  // 2. تحديد المواعيد الجديدة التي نحتاج لإضافتها
  newTimeSlots.forEach(newSlot => {
    const trimmedNewSlot = newSlot.toString().trim();
    if (!timesAlreadyInRowAndNewList.has(trimmedNewSlot)) {
      // الميعاد ده في القائمة الجديدة لكن مش موجود في الصف القديم، يبقى لازم نضيفه
      slotsToAdd.push(trimmedNewSlot);
    }
  });

  // 3. تنفيذ عمليات المسح
  if (rangesToClear.length > 0) {
    rangesToClear.forEach(range => range.clearContent());
    Logger.log(`[updateTeacherAvailableSchedules] Cleared ${rangesToClear.length} old slots no longer present in new list.`);
  }

  // 4. تنفيذ عمليات الإضافة (ملء الخلايا الفارغة)
  if (slotsToAdd.length > 0) {
    let addedCount = 0;
    let currentSearchCol = startColumnForTimes; // نبدأ البحث عن خانات فارغة من أول عمود للمواعيد

    while (addedCount < slotsToAdd.length) {
      // التأكد من وجود أعمدة كافية
      const currentMaxColumns = availableSchedulesSheet.getMaxColumns();
      if (currentSearchCol + 1 > currentMaxColumns) { // +1 لتحويل إلى 1-indexed
        availableSchedulesSheet.insertColumnAfter(currentMaxColumns);
        Logger.log(`[updateTeacherAvailableSchedules] Added a new column after column ${currentMaxColumns}.`);
        // بعد إضافة عمود، نحتاج لإعادة قراءة الصف لكي تكون currentRowDisplayValues محدثة
        // هذا يمكن أن يكون مكلفًا إذا تم إضافة أعمدة كثيرة جداً داخل هذه الحلقة
        // ولكن بالنسبة لعدد قليل من المواعيد هذا مقبول.
        currentRowDisplayValues.push(''); // نضيف خلية فارغة وهمية لتمثيل العمود الجديد
      }

      // التحقق من أن الخلية فارغة (حتى لو كانت موجودة من قبل ولكن تم مسحها)
      const cellValue = availableSchedulesSheet.getRange(rowToUpdate, currentSearchCol + 1).getDisplayValue();
      if (!cellValue || cellValue.toString().trim() === "") {
        // وجدنا خلية فارغة، نكتب فيها الميعاد الجديد
        availableSchedulesSheet.getRange(rowToUpdate, currentSearchCol + 1).setValue(slotsToAdd[addedCount]);
        Logger.log(`[updateTeacherAvailableSchedules] Added new slot: '${slotsToAdd[addedCount]}' to Row ${rowToUpdate}, Col ${currentSearchCol + 1}.`);
        addedCount++;
      }
      currentSearchCol++; // ننتقل للعمود التالي في كل الأحوال
    }
  } else {
    Logger.log("[updateTeacherAvailableSchedules] No new time slots to add to this day.");
  }

  return { status: "success", message: "تم تحديث المواعيد بنجاح." };
}





/**
 * تجلب ملخصاً شاملاً لبيانات الأكاديمية (الطلاب، التجديدات، المدفوعات).
 * تُستخدم بواسطة واجهة الويب لعرض لوحة الملخص.
 * @returns {string} سلسلة JSON تحتوي على كائن ملخص البيانات.
 */
function getAcademySummary() {
  Logger.log("getAcademySummary: Starting execution.");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAME);
  const historySheet = spreadsheet.getSheetByName(HISTORY_SHEET_NAME); // لبيانات المدفوعات

  if (!masterSheet) {
    Logger.log(`خطأ: شيت المشرف الرئيسي '${MASTER_SHEET_NAME}' غير موجود.`);
    return JSON.stringify({});
  }
  if (!historySheet) {
    Logger.log(`خطأ: شيت السجل التاريخي '${HISTORY_SHEET_NAME}' غير موجود.`);
    // يمكن الاستمرار إذا لم تكن بيانات المدفوعات حاسمة، أو إرجاع خطأ.
  }

  let totalStudents = 0;
  let activeStudents = 0; // الطلاب الذين لديهم حصص متبقية أو حالة "تم التجديد"
  let renewalNeeded = 0; // الطلاب الذين حالتهم "مطلوب التجديد" أو فارغة
  let overdueRenewal = 0; // الطلاب الذين حالتهم "تجاوز الحد" أو حالات "إيقاف"
  let monthlyPayments = 0;
  let yearlyPayments = 0;

  // 1. جلب بيانات الطلاب من الشيت الرئيسي
  const masterData = masterSheet.getDataRange().getValues();
  // تأكد من أن MASTER_FIRST_DATA_ROW صحيحة
  const studentRows = masterData.slice(MASTER_FIRST_DATA_ROW - 1);

  totalStudents = studentRows.length; // إجمالي عدد الصفوف بعد أول صفوف

  studentRows.forEach(row => {
    const status = (row[MASTER_RENEWAL_STATUS_COL - 1] || '').toString().trim();
    const attendance = parseFloat(row[MASTER_TOTAL_ATTENDANCE_COL - 1]) || 0; // عدد الحصص الحاضرة

    if (status === 'تم التجديد') {
      activeStudents++; // يمكن تعريف النشط بأنه تم التجديد
    }

    if (status === 'مطلوب التجديد') {
      renewalNeeded++;
    } else if (status === 'تجاوز الحد') {
      overdueRenewal++;
    } else if (status === 'لا يوجد حضور' || status === '') { // اعتبرهم بحاجة للتجديد أو مراجعة
      renewalNeeded++;
    }
  });

  // 2. جلب بيانات المدفوعات من السجل التاريخي
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();
    const currentMonth = new Date().getMonth(); // 0-indexed
    const currentYear = new Date().getFullYear();

    // نحدد الأعمدة ذات الصلة بالشهر الحالي من رؤوس السجل التاريخي
    const historyHeadersRow1 = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
    const currentMonthName = MONTH_NAMES_AR[currentMonth];

    let monthStartColIndex = -1;
    for(let i=0; i < historyHeadersRow1.length; i++){
      if(historyHeadersRow1[i].toString().trim() === currentMonthName){
        monthStartColIndex = i + 1; // 1-indexed column number
        break;
      }
    }

    if (monthStartColIndex !== -1) {
      const currentMonthAmountColIndex = monthStartColIndex + HISTORY_MONTH_AMOUNT_OFFSET;
      
      // نمر على صفوف بيانات الطلاب في السجل التاريخي
      for (let i = HISTORY_FIRST_DATA_ROW - 1; i < historyData.length; i++) {
        const row = historyData[i];
        const paymentDate = row[monthStartColIndex + HISTORY_MONTH_DATE_OFFSET -1]; // عمود التاريخ الفعلي
        const amount = parseFloat(row[currentMonthAmountColIndex - 1]) || 0;

        // التأكد من أن المبلغ هو لهذا الشهر والسنة (إذا كان التاريخ متاحاً)
        if (amount > 0) {
            if (paymentDate instanceof Date && paymentDate.getMonth() === currentMonth && paymentDate.getFullYear() === currentYear) {
                monthlyPayments += amount;
            } else if (typeof paymentDate === 'string' && paymentDate.includes(currentMonthName) && paymentDate.includes(currentYear.toString())) {
                // إذا كان التاريخ نصاً (يمكن تحسين هذا التحقق)
                monthlyPayments += amount;
            }
            // إذا لم يكن التاريخ متوفراً، افترض أنه لهذا الشهر
            // monthlyPayments += amount; 
        }
        
        // لحساب الدفعات السنوية، يجب أن نمر على كل أعمدة المدفوعات لكل شهر في تاريخ هذا العام
        // هذا يتطلب منطقاً أكثر تعقيداً لتحديد جميع أعمدة المدفوعات لجميع الشهور في هذا العام
        // أو دمج بيانات الدفع في عمود واحد لسهولة التجميع
        // حالياً، سنترك yearlyPayments كافتراضي أو نقوم بجمع كل المبالغ الموجودة لتوضيح الفكرة.
        // هذا الجزء يحتاج إلى تعديل دقيق بناءً على هيكل شيت السجل التاريخي.
        for (let col = HISTORY_FIRST_MONTH_COL_START -1; col < row.length; col += HISTORY_MONTH_TRIPLET_SIZE) {
            const historyPaymentAmount = parseFloat(row[col + HISTORY_MONTH_AMOUNT_OFFSET - 1]) || 0;
            // يمكن إضافة منطق للتحقق من سنة المدفوعات هنا إذا كانت الأعمدة تتكرر لسنوات
            yearlyPayments += historyPaymentAmount;
        }

      }
    }
  }


  // تجهيز كائن الملخص
  const summary = {
    totalStudents: totalStudents,
    activeStudents: activeStudents, // تحتاج لتعريف دقيق لما هو "طالب نشط"
    renewalNeeded: renewalNeeded,
    overdueRenewal: overdueRenewal,
    monthlyPayments: monthlyPayments,
    yearlyPayments: yearlyPayments // تحتاج لمنطق جلب دقيق
  };

  Logger.log("getAcademySummary: Summary data:", JSON.stringify(summary));
  return JSON.stringify(summary);
}
