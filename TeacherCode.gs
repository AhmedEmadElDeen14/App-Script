/**
 * هذا الملف يحتوي على دوال App Script الخاصة بشيت المعلم.
 * يستخدمها المعلم لتسجيل حضور وغياب الطلاب ولتسجيل حضوره الخاص.
 */

// ==============================================================================
// 1. الدوال الأساسية لنقطة الدخول والوظائف المساعدة العامة في شيت المعلم
// ==============================================================================


const SUPERVISOR_SHEET_ID = '1ujHL2gFuEQzbS4-9KusLgm5ObOzRsvxZnjah8CCfgsU';




/**
 * تُسجل الدقائق المحتسبة للمعلم في شيت "سجل حضور المعلم" الخاص به.
 *
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} teacherName - اسم المعلم.
 * @param {number} minutesToAdd - عدد الدقائق لإضافتها.
 * @param {Date} date - تاريخ الحصة.
 * @param {string} monthYear - الشهر والسنة بتنسيق YYYY-MM.
 */
function recordTeacherMinutes(teacherId, teacherName, minutesToAdd, date, monthYear) {
    const teacherPersonalAttendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل حضور المعلم");
    if (!teacherPersonalAttendanceSheet) {
        Logger.log("خطأ: شيت 'سجل حضور المعلم' غير موجود في ملف المعلم لـ recordTeacherMinutes.");
        return;
    }

    // التأكد من أن الدقائق قيمة صالحة
    if (typeof minutesToAdd !== 'number' || minutesToAdd <= 0) {
        Logger.log(`تحذير: تم استدعاء recordTeacherMinutes بقيمة دقائق غير صالحة: ${minutesToAdd}`);
        return;
    }

    let personalAttendanceRowIndex = -1;
    const personalAttendanceData = teacherPersonalAttendanceSheet.getDataRange().getValues();

    // البحث عن سجل حضور المعلم لهذا الشهر والسنة
    for (let i = 1; i < personalAttendanceData.length; i++) {
        const row = personalAttendanceData[i];
        const logTeacherId = String(row[0] || '').trim(); // العمود A: Teacher ID
        const logMonthYear = String(row[1] || '').trim(); // العمود B: Month (YYYY-MM)

        if (logTeacherId === teacherId && logMonthYear === monthYear) {
            personalAttendanceRowIndex = i;
            break;
        }
    }

    if (personalAttendanceRowIndex !== -1) {
        // تحديث السجل الموجود: إضافة الدقائق إلى عمود Total Session Minutes (العمود D)
        let currentTotalMinutes = teacherPersonalAttendanceSheet.getRange(personalAttendanceRowIndex + 1, 4).getValue(); // العمود D
        currentTotalMinutes = (typeof currentTotalMinutes === 'number') ? currentTotalMinutes : 0;
        teacherPersonalAttendanceSheet.getRange(personalAttendanceRowIndex + 1, 4).setValue(currentTotalMinutes + minutesToAdd);

        // يمكن تحديث عمود Status إلى "حاضر" إذا كان قد تم تعيينه لشيء آخر
        const currentStatus = String(teacherPersonalAttendanceSheet.getRange(personalAttendanceRowIndex + 1, 3).getValue() || '').trim();
        if (currentStatus === "" || currentStatus === "غائب" || currentStatus === "إجازة") {
             teacherPersonalAttendanceSheet.getRange(personalAttendanceRowIndex + 1, 3).setValue("حاضر"); // العمود C: Status
        }

        // تحديث تاريخ آخر تسجيل (العمود E) إذا لزم الأمر
        teacherPersonalAttendanceSheet.getRange(personalAttendanceRowIndex + 1, 5).setValue(date); // العمود E: Date
        Logger.log(`تم تحديث إجمالي دقائق حضور المعلم ${teacherName} إلى ${currentTotalMinutes + minutesToAdd} لهذا الشهر.`);
    } else {
        // إنشاء سجل جديد لحضور المعلم لهذا الشهر
        // الأعمدة في شيت "سجل حضور المعلم": Teacher ID (A), Month (B), Status (C), Total Session Minutes (D), Date (E), Notes (F)
        teacherPersonalAttendanceSheet.appendRow([
            teacherId,
            monthYear,
            "حاضر", // الحالة الافتراضية
            minutesToAdd, // مجموع الدقائق للحصة الأولى
            date,
            "أول حصة مسجلة لهذا الشهر"
        ]);
        Logger.log(`تم إنشاء سجل جديد لحضور المعلم ${teacherName} لهذا الشهر بأول ${minutesToAdd} دقيقة.`);
    }
}



/**
 * دالة نقطة الدخول الرئيسية لتطبيق الويب في شيت المعلم.
 * تقوم بعرض ملف HTML الرئيسي (TeacherUI.html).
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('TeacherUI')
      .setTitle('واجهة المعلم - أكاديمية غيث');
}

/**
 * دالة مساعدة لجلب ID المعلم واسمه ورقم هاتفه من شيت "المعلمين" في ملف المشرف.
 * هذه الدالة ستتطلب الوصول إلى ملف المشرف.
 *
 * @param {string} teacherPhone - رقم هاتف المعلم.
 * @returns {Object|null} كائن يحتوي على { teacherId, teacherName, phone } أو null.
 */
function getTeacherDetailsByPhoneFromSupervisor(teacherPhone) {

  try {
    const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
    const teachersSheet = supervisorSpreadsheet.getSheetByName("المعلمين");

    if (!teachersSheet) {
      Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين' في ملف المشرف.");
      return { error: "شيت 'المعلمين' غير موجود في ملف المشرف." };
    }

    const data = teachersSheet.getDataRange().getValues();
    const formattedPhone = String(teacherPhone).trim();
    const cleanedSearchPhone = formattedPhone.startsWith("0") ? formattedPhone.substring(1) : formattedPhone;

    for (let i = 1; i < data.length; i++) {
      const teacherId = String(data[i][0] || '').trim();
      const teacherName = String(data[i][1] || '').trim();
      const storedPhone = String(data[i][2] || '').trim(); // العمود C في شيت المعلمين بالمشرف

      const cleanedStoredPhone = storedPhone.startsWith("0") ? storedPhone.substring(1) : storedPhone;

      if (cleanedStoredPhone === cleanedSearchPhone) {
        return { teacherId: teacherId, teacherName: teacherName, phone: storedPhone };
      }
    }
    return { error: "لم يتم العثور على معلم بهذا الرقم في ملف المشرف." };
  } catch (e) {
    Logger.log("خطأ في getTeacherDetailsByPhoneFromSupervisor: " + e.message);
    return { error: `خطأ في جلب بيانات المعلم من المشرف: ${e.message}` };
  }
}

/**
 * تجلب جدول حصص الطالب لهذا المعلم في اليوم المحدد.
 * (تعتمد على شيت "المواعيد المتاحة للمعلمين" في ملف المشرف).
 *
 * @param {string} teacherId - Teacher ID لهذا المعلم.
 * @param {string} day - اليوم (مثلاً: "الأحد").
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب (Student ID, name, timeSlot).
 */
function getTeacherScheduleForDay(teacherId, day) {

  try {
    const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
    const teachersAvailableSlotsSheet = supervisorSpreadsheet.getSheetByName("المواعيد المتاحة للمعلمين");
    const studentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب");
    const trialStudentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب التجريبيون");


    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' في ملف المشرف غير موجود.");
    if (!studentsSheet) throw new Error("شيت 'الطلاب' في ملف المشرف غير موجود.");
    if (!trialStudentsSheet) throw new Error("شيت 'الطلاب التجريبيون' في ملف المشرف غير موجود.");

    // جلب بيانات الطلاب (ID -> Name) من كلا الشيتين في ملف المشرف
    const studentIdToNameMap = new Map();
    studentsSheet.getDataRange().getValues().forEach(row => {
      const id = String(row[0] || '').trim();
      const name = String(row[1] || '').trim();
      if (id) studentIdToNameMap.set(id, name);
    });
    trialStudentsSheet.getDataRange().getValues().forEach(row => {
      const id = String(row[0] || '').trim();
      const name = String(row[1] || '').trim();
      if (id) studentIdToNameMap.set(id, name);
    });

    const studentsForSchedule = [];
    const teacherSlotsData = teachersAvailableSlotsSheet.getDataRange().getValues();
    const headers = teacherSlotsData[0];
    const timeSlotCols = [];
    const startColIndexForSlots = 2; // العمود C

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

          // لو الخلية تحتوي على Student ID أو Trial ID (محجوزة لطالب)
          if (slotValue.startsWith("STD") || slotValue.startsWith("TRL") || slotValue.startsWith("p ")) { // 'p ' for old placeholder values
            const studentIdInCell = slotValue;
            const studentName = studentIdToNameMap.get(studentIdInCell) || studentIdInCell;

            studentsForSchedule.push({
              studentID: studentIdInCell,
              name: studentName,
              timeSlot: timeSlotHeader // الميعاد برأس العمود (HH:mm - HH:mm)
            });
          }
        });
      }
    }
    return studentsForSchedule;
  } catch (e) {
    Logger.log("خطأ في getTeacherScheduleForDay: " + e.message);
    return { error: `خطأ في جلب جدول الحصص: ${e.message}` };
  }
}

/**
 * دالة لتوليد معرف حضور جديد وفريد بناءً على آخر معرف في شيت "سجل الحضور" في ملف المشرف.
 * تُستخدم هذه الدالة من ملف المعلم لتسجيل الحضور في سجل المشرف.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} supervisorAttendanceSheet - شيت "سجل الحضور" في ملف المشرف.
 * @returns {string} معرف الحضور الجديد (مثال: ATT001).
 */
function generateUniqueAttendanceIdInSupervisor(supervisorAttendanceSheet) {
  const lastRow = supervisorAttendanceSheet.getLastRow();
  let lastGeneratedIdNum = 0;
  if (lastRow >= 2) {
    const attIds = supervisorAttendanceSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // العمود A لـ Attendance ID
    const numericIds = attIds.flat().map(id => {
      const numPart = String(id).replace('ATT', '');
      return parseInt(numPart) || 0;
    }).filter(Number);
    lastGeneratedIdNum = numericIds.length > 0 ? Math.max(...numericIds) : 0;
  }
  return `ATT${(lastGeneratedIdNum + 1).toString().padStart(3, '0')}`;
}

/**
 * تسجل حضور/غياب/تأجيل الطالب في شيت "جدول الحصص والتحضير" الخاص بالمعلم.
 * ثم تُرسل هذا التحديث إلى شيت المشرف وتُحدّث سجل الحضور المركزي وحالة الاشتراكات.
 * وإذا كان التسجيل "حضر" أو "غاب" (مخصوم)، يتم احتساب حضور المعلم في شيت "سجل حضور المعلم".
 *
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} studentId - Student ID أو Trial ID للطالب.
 * @param {string} studentName - اسم الطالب.
 * @param {string} day - يوم الحصة.
 * @param {string} timeSlot - ميعاد الحصة.
 * @param {string} status - حالة الحضور ("حضر" أو "غاب" أو "تأجيل").
 * @returns {Object} رسالة نجاح أو خطأ.
 */
function recordStudentAttendanceInTeacherSheet(teacherId, studentId, studentName, day, timeSlot, status) {
  const teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("جدول الحصص والتحضير");
  const teacherPersonalAttendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل حضور المعلم");
  if (!teacherSheet) throw new Error("شيت 'جدول الحصص والتحضير' غير موجود في ملف المعلم.");
  if (!teacherPersonalAttendanceSheet) throw new Error("شيت 'سجل حضور المعلم' غير موجود في ملف المعلم.");

  const today = new Date();
  const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const currentMonthYear = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM");
  const lastUpdatedBy = Session.getActiveUser().getEmail();
  const lock = LockService.getScriptLock();

  // SUPERVISOR_SHEET_ID هو متغير عام تم تعريفه في أعلى الملف.

  try {
    lock.waitLock(30000);

    if (SUPERVISOR_SHEET_ID === 'YOUR_SUPERVISOR_SHEET_ID_HERE' || !SUPERVISOR_SHEET_ID) {
        throw new Error("لم يتم تحديد معرف ملف المشرف في كود المعلم.");
    }

    const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
    const supervisorAttendanceLogSheet = supervisorSpreadsheet.getSheetByName("سجل الحضور");
    const supervisorSubscriptionsSheet = supervisorSpreadsheet.getSheetByName("الاشتراكات الحالية");
    const supervisorPackagesSheet = supervisorSpreadsheet.getSheetByName("الباقات");


    if (!supervisorAttendanceLogSheet) throw new Error("شيت 'سجل الحضور' غير موجود في ملف المشرف.");
    if (!supervisorSubscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود في ملف المشرف.");
    if (!supervisorPackagesSheet) throw new Error("شيت 'الباقات' غير موجود في ملف المشرف.");


    // 1. تسجيل الحضور/الغياب/التأجيل في شيت "جدول الحصص والتحضير" الخاص بالمعلم
    // الأعمدة في شيت "جدول الحصص والتحضير": Teacher ID, Day, Time Slot, Student ID, Student Name, Status, Date, Notes, Last Updated By
    teacherSheet.appendRow([
      teacherId,
      day,
      timeSlot,
      studentId,
      studentName,
      status,
      todayFormatted,
      "", // ملاحظات
      lastUpdatedBy
    ]);
    Logger.log(`تم تسجيل ${status} الطالب ${studentName} في ملف المعلم.`);

    // 2. تحديث سجل المشرف المركزي
    // التحقق من عدم تسجيل الحضور/الغياب مسبقًا لنفس الطالب في نفس التاريخ والميعاد في سجل المشرف
    const attendanceLogDataSupervisor = supervisorAttendanceLogSheet.getDataRange().getValues();
    for (let i = 1; i < attendanceLogDataSupervisor.length; i++) {
        const logRow = attendanceLogDataSupervisor[i];
        const logStudentID = String(logRow[1] || '').trim();
        const logTeacherID = String(logRow[2] || '').trim();
        const logDateValue = logRow[4];
        const logDate = (logDateValue instanceof Date) ? Utilities.formatDate(logDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        const logTimeSlot = String(logRow[5] || '').trim();
        const logStatus = String(logRow[7] || '').trim();

        // إذا كان السجل موجوداً بالفعل بنفس الحالة، نمنع التكرار
        if (logStudentID === studentId && logTeacherID === teacherId && logDate === todayFormatted && logTimeSlot === timeSlot && logStatus === status) {
            return { error: `تم تسجيل ${status} لهذا الطالب في هذا الميعاد وهذا اليوم مسبقًا في سجل المشرف.` };
        }
    }

    // تحديد نوع الحصة (عادية / تجريبية)
    let classType = (studentId.startsWith("TRL")) ? "تجريبية" : "عادية";
    let subscriptionId = '';
    let packageName = '';
    let totalPackageSessions = 0;
    let subscriptionRowIndex = -1;
    let sessionDurationMinutes = 0; // مدة الحصة بالدقائق
    let currentAbsencesCount = 0; // عدد الغيابات غير المخصومة للطالب في اشتراكه

    // **جديد: مؤشر عمود Absences Count في شيت الاشتراكات الحالية**
    const ABSENCES_COUNT_COL_INDEX = 13; // العمود M (إذا وضعته بعد المبلغ المتبقي L)
                                         // تأكد من هذا المؤشر بناءً على مكان العمود الجديد

    // جلب بيانات الاشتراك فقط للطلاب المشتركين لتحديثهم
    if (classType === "عادية") {
        const subscriptionsData = supervisorSubscriptionsSheet.getDataRange().getValues();
        for (let i = 1; i < subscriptionsData.length; i++) {
            if (String(subscriptionsData[i][1] || '').trim() === studentId) { // العمود B: Student ID
                subscriptionRowIndex = i;
                subscriptionId = String(subscriptionsData[i][0] || '').trim(); // Subscription ID
                packageName = String(subscriptionsData[i][2] || '').trim(); // اسم الباقة
                currentAbsencesCount = subscriptionsData[i][ABSENCES_COUNT_COL_INDEX - 1] || 0; // قراءة عدد الغيابات. المؤشر -1 لأنه 0-based
                
                // جلب عدد الحصص الكلي ومدة الحصة من شيت الباقات في ملف المشرف
                const packagesData = supervisorPackagesSheet.getDataRange().getValues();
                const packageRow = packagesData.find(pRow => String(pRow[0] || '').trim() === packageName);
                if (packageRow) {
                    totalPackageSessions = packageRow[2]; // العمود C: عدد الحصص الكلي
                    sessionDurationMinutes = packageRow[1]; // العمود B: مدة الحصة (دقيقة)
                }
                break;
            }
        }
    } else { // لو طالب تجريبي
        sessionDurationMinutes = 30; // افتراضياً، حصة تجريبية 30 دقيقة
    }

    // تسجيل الحضور/الغياب/التأجيل في شيت "سجل الحضور" المركزي
    const attendanceId = generateUniqueAttendanceIdInSupervisor(supervisorAttendanceLogSheet);
    supervisorAttendanceLogSheet.appendRow([
      attendanceId,
      studentId,
      teacherId,
      subscriptionId,
      today,
      timeSlot,
      day,
      status, // <--- الحالة هنا ستكون "حضر" أو "غاب" أو "تأجيل"
      classType,
      `تم التسجيل بواسطة المعلم: ${Session.getActiveUser().getEmail()}`
    ]);
    Logger.log(`تم تسجيل ${status} للطالب ${studentName} (ID: ${studentId}) في سجل المشرف مباشرة من ملف المعلم.`);

    // 3. تطبيق المنطق الجديد على حصص الطالب والمعلم
    if (classType === "عادية" && subscriptionRowIndex !== -1) { // فقط للطلاب المشتركين
        if (status === "حضر") {
            // أ. احتساب الحصة للطالب
            const currentAttendedSessionsCell = supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, 7); // العمود G
            let currentSessions = currentAttendedSessionsCell.getValue();
            currentSessions = (typeof currentSessions === 'number') ? currentSessions : 0;
            supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, 7).setValue(currentSessions + 1); // زيادة عدد الحصص الحاضرة
            Logger.log(`تم تحديث عدد الحصص الحاضرة للطالب ${studentId} إلى ${currentSessions + 1} في سجل المشرف.`);

            // تحديث "الحالة التفصيلية للتجديد"
            if (totalPackageSessions > 0 && (currentSessions + 1) >= totalPackageSessions) {
                supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue("يحتاج للتجديد"); // العمود H
                Logger.log(`حالة تجديد الطالب ${studentId} تم تحديثها إلى "يحتاج للتجديد".`);
            }

            // ب. احتساب الحصة للمعلم (دقائق)
            recordTeacherMinutes(teacherId, studentName, sessionDurationMinutes, today, currentMonthYear); // دالة جديدة
        } else if (status === "غاب") {
            // ج. منطق الغياب: إذا كان الغياب الأول، لا يُخصم، وإلا يُخصم.
            if (currentAbsencesCount < 1) { // أول غياب (غير مخصوم)
                supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, ABSENCES_COUNT_COL_INDEX).setValue(currentAbsencesCount + 1); // زيادة عدد الغيابات غير المخصومة
                Logger.log(`أول غياب للطالب ${studentId}، لم يُخصم من الحصص وتم زيادة عداد الغيابات غير المخصومة.`);
                // لا يتم احتسابها للمعلم
            } else { // الغياب الثاني أو أكثر (مخصوم)
                // أ. احتساب الحصة للطالب (خصم)
                const currentAttendedSessionsCell = supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, 7); // العمود G
                let currentSessions = currentAttendedSessionsCell.getValue();
                currentSessions = (typeof currentSessions === 'number') ? currentSessions : 0;
                supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, 7).setValue(currentSessions + 1); // خصم الحصة بزيادة عدد الحاضر
                Logger.log(`الغياب الثاني أو أكثر للطالب ${studentId}، تم خصمه من الحصص وزيادة عدد الحاضر.`);

                // تحديث "الحالة التفصيلية للتجديد"
                if (totalPackageSessions > 0 && (currentSessions + 1) >= totalPackageSessions) {
                    supervisorSubscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue("يحتاج للتجديد");
                }

                // ب. احتساب الحصة للمعلم (دقائق)
                recordTeacherMinutes(teacherId, studentName, sessionDurationMinutes, today, currentMonthYear); // يتم احتساب الغياب المخصوم للمعلم
            }
        }
        // د. لا يتم احتساب "تأجيل" للطالب أو المعلم.
        // لا نحتاج لـ else if (status === "تأجيل") هنا، لأنه لا يوجد خصم أو احتساب.
    } else if (classType === "تجريبية" && status === "حضر") {
        // إذا كان طالباً تجريبياً وحضر، يتم احتسابها للمعلم
        recordTeacherMinutes(teacherId, studentName, sessionDurationMinutes, today, currentMonthYear);
    }

    return { success: `تم تسجيل ${status} الطالب ${studentName} في ملف المعلم وتحديث سجل المشرف.` };

  } catch (e) {
    Logger.log("خطأ في recordStudentAttendanceInTeacherSheet: " + e.message);
    return { error: `فشل تسجيل الحضور/الغياب وتحديث سجل المشرف: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


// ==============================================================================
// 3. الدوال الخاصة بصفحة طلاب المعلم (عرض كل الطلاب والمواعيد)
// ==============================================================================

/**
 * تجلب بيانات جميع الطلاب (المشتركين والتجريبيين) المرتبطين بمعلم محدد.
 * تُستخدم لصفحة "طلابي ومواعيدهم" في واجهة المعلم.
 *
 * @param {string} teacherId - Teacher ID للمعلم المراد جلب طلابه.
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب الموحدة.
 * أو {Object} كائن خطأ.
 */
function getAllStudentsForTeacher(teacherId) {
  const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
  const studentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب");
  const subscriptionsSheet = supervisorSpreadsheet.getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = supervisorSpreadsheet.getSheetByName("المواعيد المتاحة للمعلمين");
  const trialStudentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب التجريبيون");

  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود في ملف المشرف." };
  if (!subscriptionsSheet) return { error: "شيت 'الاشتراكات الحالية' غير موجود في ملف المشرف." };
  if (!teachersAvailableSlotsSheet) return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود في ملف المشرف." };
  if (!trialStudentsSheet) return { error: "شيت 'الطلاب التجريبيون' غير موجود في ملف المشرف." };

  const allTeacherStudents = [];

  // جلب بيانات المواعيد المحجوزة لكل طالب (Student ID/Trial ID -> [{day, timeSlotHeader}])
  const studentBookedSlotsMap = new Map();
  const availableSlotsValues = teachersAvailableSlotsSheet.getDataRange().getValues();
  const headers = availableSlotsValues[0];
  const timeSlotHeaders = [];
  const startColIndexForSlots = 2;
  for (let i = startColIndexForSlots; i < headers.length; i++) {
      const header = String(headers[i] || '').trim();
      if (header) {
          timeSlotHeaders.push({ index: i, header: header });
      }
  }

  for (let i = 1; i < availableSlotsValues.length; i++) {
      const row = availableSlotsValues[i];
      const teacherIdInSlot = String(row[0] || '').trim();
      const dayOfWeek = String(row[1] || '').trim();

      if (teacherIdInSlot === teacherId) { // فقط المواعيد الخاصة بهذا المعلم
          timeSlotHeaders.forEach(colInfo => {
              const slotValue = String(row[colInfo.index] || '').trim();
              const timeSlotHeader = colInfo.header;

              if (slotValue.startsWith("STD") || slotValue.startsWith("TRL") || slotValue.startsWith("p ")) {
                  const studentIdInCell = slotValue;
                  if (!studentBookedSlotsMap.has(studentIdInCell)) {
                      studentBookedSlotsMap.set(studentIdInCell, []);
                  }
                  studentBookedSlotsMap.get(studentIdInCell).push({
                      day: dayOfWeek,
                      timeSlotHeader: timeSlotHeader,
                      teacherId: teacherIdInSlot
                  });
              }
          });
      }
  }

  // جلب بيانات الاشتراكات (Student ID -> Subscription Details)
  const subscriptionsMap = new Map();
  const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
  subscriptionsData.forEach((row, index) => {
    if (index === 0) return;
    const studentID = String(row[1] || '').trim();
    const subTeacherId = String(row[3] || '').trim(); // Teacher ID في الاشتراك
    if (studentID && subTeacherId === teacherId) { // فقط الاشتراكات لهذا المعلم
      subscriptionsMap.set(studentID, {
        packageName: String(row[2] || '').trim(),
        renewalStatus: String(row[7] || '').trim(),
        // يمكن إضافة المزيد من تفاصيل الاشتراك هنا حسب الحاجة
      });
    }
  });


  // 1. معالجة الطلاب المشتركين لهذا المعلم
  const studentsData = studentsSheet.getDataRange().getValues();
  studentsData.forEach((row, index) => {
    if (index === 0) return;
    const studentID = String(row[0] || '').trim();

    const subscriptionDetails = subscriptionsMap.get(studentID);
    if (subscriptionDetails) { // إذا كان الطالب مشتركًا لهذا المعلم
      const studentInfo = {
        studentID: studentID,
        name: String(row[1] || '').trim(),
        age: row[2],
        phone: String(row[3] || '').trim(),
        basicStatus: String(row[7] || '').trim(), // الحالة الأساسية للطالب
        packageName: subscriptionDetails.packageName,
        renewalStatus: subscriptionDetails.renewalStatus,
        //Teacher Name (هذا هو المعلم الحالي)
      };

      const bookedSlots = studentBookedSlotsMap.get(studentID) || [];
      studentInfo.day1 = bookedSlots[0] ? bookedSlots[0].day : '';
      studentInfo.time1 = bookedSlots[0] ? bookedSlots[0].timeSlotHeader : '';
      studentInfo.day2 = bookedSlots[1] ? bookedSlots[1].day : '';
      studentInfo.time2 = bookedSlots[1] ? bookedSlots[1].timeSlotHeader : '';
      
      allTeacherStudents.push(studentInfo);
    }
  });

  // 2. معالجة الطلاب التجريبيين لهذا المعلم
  const trialStudentsRawData = trialStudentsSheet.getDataRange().getValues();
  trialStudentsRawData.forEach((row, index) => {
    if (index === 0) return;
    const trialID = String(row[0] || '').trim();
    const trialTeacherId = String(row[4] || '').trim(); // Teacher ID في شيت التجريبيين

    if (trialTeacherId === teacherId) { // إذا كان الطالب التجريبي لهذا المعلم
        const trialStudentInfo = {
            trialID: trialID,
            name: String(row[1] || '').trim(),
            age: row[2],
            phone: String(row[3] || '').trim(),
            basicStatus: String(row[10] || '').trim(), // الحالة من عمود Status في شيت التجريبيين
            packageName: 'تجريبي', // الباقة افتراضياً تجريبي
            renewalStatus: 'تجريبي', // حالة التجديد افتراضياً تجريبي
            //Teacher Name (هذا هو المعلم الحالي)
        };
        // المواعيد المباشرة من شيت الطلاب التجريبيين
        trialStudentInfo.day1 = String(row[6] || '').trim();
        trialStudentInfo.time1 = String(row[7] || '').trim();
        trialStudentInfo.day2 = ''; // الطلاب التجريبيين لهم ميعاد واحد
        trialStudentInfo.time2 = '';

        allTeacherStudents.push(trialStudentInfo);
    }
  });

  Logger.log(`تم جلب ${allTeacherStudents.length} طالب للمعلم ID ${teacherId}.`);
  return allTeacherStudents;
}

/**
 * تجلب قائمة باقات الاشتراك من شيت "الباقات" في ملف المشرف.
 * تُستخدم لفلتر الباقات في واجهة المعلم.
 * @returns {Array<string>} قائمة بأسماء الباقات.
 */
function getSubscriptionPackageListFromSupervisor() {
  const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
  const packagesSheet = supervisorSpreadsheet.getSheetByName("الباقات");

  if (!packagesSheet) {
    Logger.log("خطأ: لم يتم العثور على شيت 'الباقات' في ملف المشرف.");
    return { error: "شيت 'الباقات' غير موجود في ملف المشرف." };
  }
  const data = packagesSheet.getDataRange().getValues();
  const packageNames = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim(); // العمود A: اسم الباقة
    if (name) {
      packageNames.push(name);
    }
  }
  return packageNames;
}

// ==============================================================================
// 4. الدوال الخاصة بصفحة إدارة مواعيد المعلم (في ملف المعلم - تتصل بالمشرف)
// ==============================================================================

/**
 * تجلب كافة المواعيد (المتاحة والمحجوزة) لمعلم معين من شيت "المواعيد المتاحة للمعلمين" في ملف المشرف.
 * تُستخدم لصفحة "إدارة مواعيدي" في واجهة المعلم.
 *
 * @param {string} teacherId - Teacher ID للمعلم المراد جلب مواعيده.
 * @returns {Object} كائن يحتوي على اسم المعلم، Teacher ID، ومصفوفة من كائنات المواعيد، أو كائن خطأ.
 */
function getTeacherAvailableSlotsFromSupervisor(teacherId) {
  const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
  const teachersSheet = supervisorSpreadsheet.getSheetByName("المعلمين");
  const teachersAvailableSlotsSheet = supervisorSpreadsheet.getSheetByName("المواعيد المتاحة للمعلمين");
  const studentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب");
  const trialStudentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب التجريبيون");

  if (!teachersSheet) return { error: "شيت 'المعلمين' غير موجود في ملف المشرف." };
  if (!teachersAvailableSlotsSheet) return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود في ملف المشرف." };
  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود في ملف المشرف." };
  if (!trialStudentsSheet) return { error: "شيت 'الطلاب التجريبيون' غير موجود في ملف المشرف." };

  const teacherName = getTeacherNameByIdFromSupervisor(teacherId); // دالة مساعدة (قد تحتاج لإضافتها إذا لم تكن موجودة)
  if (!teacherName) {
    return { error: `لم يتم العثور على اسم المعلم لـ Teacher ID: ${teacherId}` };
  }

  const slots = [];
  const timeSlotColumns = [];
  const startColIndexForSlots = 2; // العمود C (مؤشر 2)

  // جلب أسماء الطلاب (STD/TRL ID -> Name) من ملف المشرف
  const studentIdToNameMap = new Map();
  studentsSheet.getDataRange().getValues().forEach(row => {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (id) studentIdToNameMap.set(id, name);
  });
  trialStudentsSheet.getDataRange().getValues().forEach(row => {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (id) studentIdToNameMap.set(id, name);
  });


  const headers = teachersAvailableSlotsSheet.getDataRange().getValues()[0];
  for (let i = startColIndexForSlots; i < headers.length; i++) {
    const header = String(headers[i] || '').trim();
    if (header) {
      timeSlotColumns.push({ index: i, header: header });
    }
  }

  const availableSlotsData = teachersAvailableSlotsSheet.getDataRange().getValues();
  // البحث عن صفوف المعلم المحددة في شيت المواعيد
  for (let i = 1; i < availableSlotsData.length; i++) {
    const row = availableSlotsData[i];
    const currentTeacherIdInSheet = String(row[0] || '').trim();
    const dayOfWeek = String(row[1] || '').trim();

    if (currentTeacherIdInSheet === teacherId) { // وجدنا صفوف هذا المعلم
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
          timeSlotHeader: timeSlotHeader,
          actualSlotValue: slotValue,
          isBooked: isBooked,
          bookedBy: bookedBy
        });
      });
    }
  }

  Logger.log(`تم جلب ${slots.length} موعد للمعلم ID ${teacherId} من ملف المشرف.`);
  return { teacherName: teacherName, teacherId: teacherId, slots: slots };
}

/**
 * دالة مساعدة: تجلب اسم المعلم من Teacher ID (العمود A) من شيت "المعلمين" في ملف المشرف.
 *
 * @param {string} teacherId - Teacher ID.
 * @returns {string|null} اسم المعلم (من العمود B) أو null.
 */
function getTeacherNameByIdFromSupervisor(teacherId) {
    const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
    const teachersSheet = supervisorSpreadsheet.getSheetByName("المعلمين");
    if (!teachersSheet) {
        Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين' في ملف المشرف لـ getTeacherNameByIdFromSupervisor.");
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
 * تقوم بتحديث المواعيد المتاحة لمعلم في يوم معين في شيت "المواعيد المتاحة للمعلمين" في ملف المشرف.
 * يتم وضع رأس العمود في الخلايا المختارة (لجعلها متاحة) وإفراغ الخلايا غير المختارة (لجعلها غير متاحة).
 * لا يمكن تعديل المواعيد المحجوزة بالفعل.
 *
 * @param {string} teacherId - Teacher ID للمعلم.
 * @param {string} day - اليوم الذي يتم تحديث مواعيده.
 * @param {Array<Object>} selectedSlots - مصفوفة من كائنات { timeSlotHeader: "HH:MM - HH:MM" } تمثل المواعيد التي يجب أن تكون متاحة.
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateTeacherSlotsInSupervisor(teacherId, day, selectedSlots) {
  const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
  const teachersAvailableSlotsSheet = supervisorSpreadsheet.getSheetByName("المواعيد المتاحة للمعلمين");
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' غير موجود في ملف المشرف.");

    const data = teachersAvailableSlotsSheet.getDataRange().getValues();
    const headers = data[0];

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
      throw new Error(`لم يتم العثور على صف مواعيد للمعلم ID ${teacherId} في يوم ${day} في ملف المشرف.`);
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
      if (currentCellValue.startsWith("STD") || currentCellValue.startsWith("TRL") || (currentCellValue !== '' && currentCellValue !== timeSlotHeader)) {
        Logger.log(`الموعد ${day} ${timeSlotHeader} محجوز حالياً (${currentCellValue}). لن يتم تعديله في ملف المشرف.`);
        return; // انتقل إلى الموعد التالي
      }

      // الحالة: الموعد يجب أن يصبح متاحاً
      if (isSelectedInModal && currentCellValue !== timeSlotHeader) {
        updates.push({
          row: teacherRowIndex + 1,
          col: colIndex + 1,
          value: timeSlotHeader // اجعلها متاحة بوضع رأس العمود
        });
        Logger.log(`جعل الموعد ${day} ${timeSlotHeader} متاحاً في ملف المشرف.`);
      }
      // الحالة: الموعد يجب أن يصبح غير متاح (فارغ)
      else if (!isSelectedInModal && currentCellValue === timeSlotHeader) {
        updates.push({
          row: teacherRowIndex + 1,
          col: colIndex + 1,
          value: '' // اجعلها فارغة (غير متاحة)
        });
        Logger.log(`جعل الموعد ${day} ${timeSlotHeader} غير متاح في ملف المشرف.`);
      }
    });

    // تطبيق جميع التحديثات
    if (updates.length > 0) {
      updates.forEach(update => {
          teachersAvailableSlotsSheet.getRange(update.row, update.col).setValue(update.value);
      });
      Logger.log(`تم تطبيق ${updates.length} تحديث على مواعيد المعلم ID ${teacherId} في يوم ${day} بملف المشرف.`);
    } else {
      Logger.log("لا توجد تغييرات لتطبيقها على مواعيد المعلم بملف المشرف.");
    }

    return { success: "تم تحديث مواعيدك في ملف المشرف بنجاح." };

  } catch (e) {
    Logger.log("خطأ في updateTeacherSlotsInSupervisor: " + e.message);
    return { error: `فشل تحديث المواعيد في ملف المشرف: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}



/**
 * تجلب وتحسب ملخص أداء المعلم للشهر الحالي، بما في ذلك إحصائيات الطلاب والمواعيد.
 * تعتمد على شيت "جدول الحصص والتحضير" و "سجل حضور المعلم" الخاص بالمعلم، وشيتات المشرف.
 *
 * @param {string} teacherId - Teacher ID للمعلم المراد جلب ملخصه.
 * @returns {Object} كائن يحتوي على الإحصائيات الشهرية للمعلم.
 * أو {Object} كائن خطأ.
 */
function getTeacherMonthlySummary(teacherId) {
  const teacherClassesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("جدول الحصص والتحضير");
  const teacherPersonalAttendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل حضور المعلم");
  const supervisorSpreadsheet = SpreadsheetApp.openById(SUPERVISOR_SHEET_ID);
  const supervisorStudentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب");
  const supervisorTrialStudentsSheet = supervisorSpreadsheet.getSheetByName("الطلاب التجريبيون");
  const supervisorTeachersAvailableSlotsSheet = supervisorSpreadsheet.getSheetByName("المواعيد المتاحة للمعلمين");
  // supervisorSubscriptionsSheet ليست ضرورية في هذه الدالة بعد الآن

  if (!teacherClassesSheet) return { error: "شيت 'جدول الحصص والتحضير' غير موجود." };
  if (!teacherPersonalAttendanceSheet) return { error: "شيت 'سجل حضور المعلم' غير موجود." };
  if (!supervisorStudentsSheet) return { error: "شيت 'الطلاب' غير موجود بملف المشرف." };
  if (!supervisorTrialStudentsSheet) return { error: "شيت 'الطلاب التجريبيون' غير موجود بملف المشرف." };
  if (!supervisorTeachersAvailableSlotsSheet) return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود بملف المشرف." };
  
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();
  const currentMonthYearFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");

  const summary = {
    completedClasses: 0,
    totalAbsences: 0,
    postponedClasses: 0,
    totalWorkingMinutes: 0,
    totalWorkingHours: 0.00,
    totalRegisteredStudents: 0,
    totalBookedSlots: 0,
    totalAvailableSlots: 0
  };

  // 1. حساب الإحصائيات من شيت "جدول الحصص والتحضير" (للحصص التي تم تسجيلها هذا الشهر)
  const classRecordsData = teacherClassesSheet.getDataRange().getValues();
  const studentAbsenceCounterMap = new Map();

  classRecordsData.forEach((row, index) => {
    if (index === 0) return;
    
    const recordDate = row[6]; // العمود G: Date
    const status = String(row[5] || '').trim(); // العمود F: Status ("حضر", "غاب", "تأجيل")

    if (recordDate instanceof Date && recordDate.getMonth() === currentMonth && recordDate.getFullYear() === currentYear) {
      if (status === "حضر") {
        summary.completedClasses++;
      } else if (status === "غاب") {
        summary.totalAbsences++;
      } else if (status === "تأجيل") {
        summary.postponedClasses++;
      }
    }
  });

  // 2. حساب إجمالي دقائق العمل من شيت "سجل حضور المعلم"
  const personalAttendanceRecords = teacherPersonalAttendanceSheet.getDataRange().getValues();
  Logger.log("Tracing minutes - currentMonthYearFormatted: " + currentMonthYearFormatted); // <--- LOG
  let foundMonthlyRecord = false; // <--- LOG

  personalAttendanceRecords.forEach((row, index) => {
    if (index === 0) return;
    
    const rawRecordMonthYear = row[1]; // القيمة الخام من العمود B
    const recordMonthYear = (rawRecordMonthYear instanceof Date) ? Utilities.formatDate(rawRecordMonthYear, Session.getScriptTimeZone(), "yyyy-MM") : String(rawRecordMonthYear || '').trim(); // تنسيقها إذا كانت Date object
    const totalMinutesForMonth = row[3]; // العمود D: Total Session Minutes

    Logger.log("Tracing minutes - Checking row " + (index + 1) + ": MonthYear=" + recordMonthYear + ", Minutes=" + totalMinutesForMonth); // <--- LOG

    if (recordMonthYear === currentMonthYearFormatted && typeof totalMinutesForMonth === 'number') {
      summary.totalWorkingMinutes = totalMinutesForMonth;
      foundMonthlyRecord = true; // <--- LOG
      Logger.log("Tracing minutes - Match found for month: " + recordMonthYear + ", Total Minutes: " + totalMinutesForMonth); // <--- LOG
    }
  });
  if (!foundMonthlyRecord) { // <--- LOG
      Logger.log("Tracing minutes - No matching monthly record found for " + currentMonthYearFormatted + " or minutes not a number."); // <--- LOG
  }

  // 3. حساب الساعات النهائية
  summary.totalWorkingHours = (summary.totalWorkingMinutes / 60).toFixed(2);

  // 4. حساب عدد الطلاب المسجلين مع المعلم والمواعيد المحجوزة/المتاحة
  const studentsWithTeacher = new Set(); // لضمان الطلاب الفريدين
  
  const allSlotsData = supervisorTeachersAvailableSlotsSheet.getDataRange().getValues();
  
  for (let i = 1; i < allSlotsData.length; i++) {
    const row = allSlotsData[i];
    const teacherIdInSlot = String(row[0] || '').trim();
    if (teacherIdInSlot === teacherId) { // فقط المواعيد الخاصة بهذا المعلم
        const headers = allSlotsData[0];
        const startColIndexForSlots = 2; // العمود C
        for (let colIndex = startColIndexForSlots; colIndex < headers.length; colIndex++) {
            const slotValue = String(row[colIndex] || '').trim();
            const timeSlotHeader = String(headers[colIndex] || '').trim();

            if (slotValue.startsWith("STD") || slotValue.startsWith("TRL")) {
                studentsWithTeacher.add(slotValue);
                summary.totalBookedSlots++;
            } else if (slotValue === timeSlotHeader) { // متاح (يحتوي على رأس العمود)
                summary.totalAvailableSlots++;
            }
        }
    }
  }

  summary.totalRegisteredStudents = studentsWithTeacher.size;

  return summary;
}


/**
 * تجلب إجمالي دقائق عمل معلم محدد لشهر محدد من شيت سجل حضوره.
 * @param {string} teacherSheetId - معرف ملف Google Sheet الخاص بالمعلم.
 * @param {string} teacherId - معرف المعلم.
 * @param {string} monthYear - الشهر والسنة بتنسيق YYYY-MM (مثلاً: "2025-06").
 * @returns {number|null} إجمالي الدقائق أو null إذا لم يتم العثور على السجل.
 */
function getTotalMinutesForTeacher(teacherSheetId, teacherId, monthYear) {
  try {
    const teacherSpreadsheet = SpreadsheetApp.openById(teacherSheetId);
    const teacherPersonalAttendanceSheet = teacherSpreadsheet.getSheetByName("سجل حضور المعلم");

    if (!teacherPersonalAttendanceSheet) {
      Logger.log(`خطأ: شيت 'سجل حضور المعلم' غير موجود في ملف المعلم ID: ${teacherSheetId}`);
      return null;
    }

    const data = teacherPersonalAttendanceSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const logTeacherId = String(row[0] || '').trim(); // العمود A
      const logMonthYear = String(row[1] || '').trim(); // العمود B
      const totalMinutes = row[3]; // العمود D

      if (logTeacherId === teacherId && logMonthYear === monthYear) {
        return typeof totalMinutes === 'number' ? totalMinutes : 0;
      }
    }
    return 0; // لم يتم العثور على سجل لهذا الشهر
  } catch (e) {
    Logger.log("خطأ في getTotalMinutesForTeacher: " + e.message);
    return null;
  }
}
