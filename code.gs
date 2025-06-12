/**
 * هذا الملف يحتوي على دوال App Script لنظام إدارة أكاديمية غيث.
 * تم بناء الكود ليتوافق مع هيكلة الشيتات الجديدة.
 */

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

  // التحقق مما إذا كانت الخلية فارغة بالفعل قبل الحجز
  if (targetCell.getValue()) {
    // لو الخلية فيها قيمة، يبقى الميعاد ده محجوز بالفعل
    return { error: `الميعاد ${day} ${timeSlotHeader} محجوز بالفعل.` };
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
            if (slotValue === '') {
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
    const name = String(data[i][0] || '').trim(); // العمود A: اسم الباقة
    if (name) {
      packageNames.push(name);
    }
  }
  return packageNames;
}

/**
 * جلب قائمة حالات الدفع المتاحة (يمكن أن تكون ثابتة أو من شيت).
 * @returns {Array<string>} قائمة بحالات الدفع.
 */
function getPaymentStatusList() {
  // هذه القائمة يمكن أن تكون ثابتة أو تجلب من شيت "حالة الدفع" لو كان موجوداً
  // بناءً على تصميم الشيتات، سنفترض أنها ثابتة في الكود.
  return [
    "تم الدفع",
    "لم يتم الدفع",
    "تم دفع جزء",
    "حلقة تجريبية",
    "لم يشترك"
  ];
}


/**
 * حفظ بيانات الطالب الجديد.
 * تقوم بحفظ البيانات في شيت "الطلاب"، وإنشاء اشتراك في شيت "الاشتراكات الحالية"،
 * وتحديث المواعيد المحجوزة في شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {Object} formData - كائن يحتوي على بيانات النموذج المرسلة من الواجهة الأمامية.
 * المتوقع من الواجهة الأمامية (من حقول التسجيل):
 * { regName, regAge, regPhone, regTeacher (اسم المعلم), regDay1, regTime1,
 * regDay2, regTime2, regPaymentStatus (حالة الدفع), regSubscriptionPackage (اسم الباقة) }
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function saveData(formData) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // التحقق من وجود الشيتات الضرورية
    if (!studentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب'.");
    if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    if (!teachersSheet) throw new Error("لم يتم العثور على شيت 'المعلمين'.");

    // 1. جلب Teacher ID من اسم المعلم
    const teacherId = getTeacherIdByName(formData.regTeacher); // دالة مساعدة
    if (!teacherId) throw new Error(`لم يتم العثور على Teacher ID للمعلم: ${formData.regTeacher}`);

    // 2. حفظ الطالب في شيت "الطلاب"
    const newStudentId = generateUniqueStudentId(studentsSheet); // دالة موجودة
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // تحديد الحالة الأساسية للطالب بناءً على حالة الدفع للحلقة التجريبية
    let studentBasicStatus = "قيد التسجيل"; // افتراضي
    if (formData.regPaymentStatus === "تم الدفع") {
        studentBasicStatus = "مشترك";
    } else if (formData.regPaymentStatus === "حلقة تجريبية") {
        studentBasicStatus = "تجريبي";
    } else if (formData.regPaymentStatus === "لم يشترك" || formData.regPaymentStatus === "تم دفع جزء") {
        studentBasicStatus = "معلق"; // أو حالة أخرى مناسبة
    }


    // ترتيب الأعمدة في شيت "الطلاب": Student ID, اسم الطالب, السن, رقم الهاتف (ولي الأمر), رقم هاتف الطالب (إن وجد), البلد, تاريخ التسجيل, الحالة الأساسية للطالب, ملاحظات
    studentsSheet.appendRow([
      newStudentId,                                   // Student ID (A)
      formData.regName,                               // اسم الطالب (B)
      formData.regAge,                                // السن (C)
      String(formData.regPhone).trim(),               // رقم الهاتف (ولي الأمر) (D)
      "",                                             // رقم هاتف الطالب (إن وجد) (E)
      "",                                             // البلد (F)
      today,                                          // تاريخ التسجيل (G)
      studentBasicStatus,                             // الحالة الأساسية للطالب (H)
      ""                                              // ملاحظات (I)
    ]);
    Logger.log(`تم حفظ الطالب ${formData.regName} (ID: ${newStudentId}) في شيت 'الطلاب'.`);

    // 3. إنشاء اشتراك في شيت "الاشتراكات الحالية"
    const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet); // دالة مساعدة
    const packageName = formData.regSubscriptionPackage;
    const packageDetails = getPackageDetails(packageName); // دالة مساعدة

    let subscriptionRenewalStatus = "تم التجديد"; // افتراضي للاشتراك الجديد
    let totalClassesAttended = 0;
    let subscriptionType = "شهري"; // افتراضي
    
    // ضبط الحالة التفصيلية للتجديد ونوع الحصة بناءً على حالة الدفع
    if (formData.regPaymentStatus === "حلقة تجريبية") {
        subscriptionRenewalStatus = "تجريبي";
        totalClassesAttended = 0; // الحصص التجريبية لا تخصم من الاشتراك الأساسي
    } else { // لو مش تجريبي، نعتبره اشتراك عادي
        if (packageDetails && packageDetails['نوع الباقة']) {
            subscriptionType = packageDetails['نوع الباقة'];
        }
    }

    // حساب تاريخ نهاية الاشتراك المتوقع (بالنسبة للاشتراكات الشهرية)
    let endDate = "";
    if (subscriptionType === "شهري" && packageDetails && packageDetails['عدد الحصص الكلي'] > 0) {
        // يمكن حساب تاريخ نهاية الاشتراك بشكل أكثر تعقيدًا بناءً على عدد الحصص
        // أو ببساطة بزيادة شهر واحد من تاريخ البداية
        const startDate = new Date(today);
        startDate.setMonth(startDate.getMonth() + 1);
        endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (subscriptionType === "نصف سنوي") {
        const startDate = new Date(today);
        startDate.setMonth(startDate.getMonth() + 6); // إضافة 6 أشهر
        endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }


    // ترتيب الأعمدة في شيت "الاشتراكات الحالية": Subscription ID, Student ID, اسم الباقة, Teacher ID, تاريخ بداية الاشتراك, تاريخ نهاية الاشتراك المتوقع, عدد الحصص الحاضرة, الحالة التفصيلية للتجديد, تاريخ آخر تجديد, مبلغ الاشتراك الكلي, المبلغ المدفوع حتى الآن, المبلغ المتبقي, ملاحظات خاصة بالاشتراك
    subscriptionsSheet.appendRow([
      newSubscriptionId,                              // Subscription ID (A)
      newStudentId,                                   // Student ID (B)
      packageName,                                    // اسم الباقة (C)
      teacherId,                                      // Teacher ID (D)
      today,                                          // تاريخ بداية الاشتراك (E)
      endDate,                                        // تاريخ نهاية الاشتراك المتوقع (F)
      totalClassesAttended,                           // عدد الحصص الحاضرة (G)
      subscriptionRenewalStatus,                      // الحالة التفصيلية للتجديد (H)
      today,                                          // تاريخ آخر تجديد (I)
      packageDetails ? packageDetails['السعر'] : 0,   // مبلغ الاشتراك الكلي (J)
      formData.regPaymentStatus === "تم الدفع" ? (packageDetails ? packageDetails['السعر'] : 0) : 0, // المبلغ المدفوع حتى الآن (K)
      formData.regPaymentStatus === "تم الدفع" ? 0 : (packageDetails ? packageDetails['السعر'] : (packageDetails ? packageDetails['السعر'] : 0)), // المبلغ المتبقي (L)
      ""                                              // ملاحظات خاصة بالاشتراك (M)
    ]);
    Logger.log(`تم إنشاء اشتراك (${newSubscriptionId}) للطالب ${newStudentId} في شيت 'الاشتراكات الحالية'.`);

    // 4. حجز المواعيد في شيت "المواعيد المتاحة للمعلمين"
    const bookedSlotsToUpdate = [];
    let bookingTypeForSlot = (formData.regPaymentStatus === "حلقة تجريبية") ? "تجريبي" : "عادي";

    // الميعاد الأول
    if (formData.regDay1 && formData.regTime1) {
        bookedSlotsToUpdate.push({
            teacherId: teacherId,
            day: formData.regDay1,
            timeSlotHeader: formData.regTime1, // رأس العمود (مثل 09:00 - 09:30)
            studentId: newStudentId,
            bookingType: bookingTypeForSlot
        });
    }
    // الميعاد الثاني
    if (formData.regDay2 && formData.regTime2) {
        bookedSlotsToUpdate.push({
            teacherId: teacherId,
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
            Logger.log(`خطأ في حجز الميعاد ${slot.day} ${slot.timeSlotHeader} للطالب ${newStudentId}: ${result.error}`);
            // إذا فشل حجز ميعاد، نلغي عملية الحفظ كلها ونرمي خطأ
            throw new Error(`تعذر حجز الميعاد ${slot.day} ${slot.timeSlotHeader}: ${result.error}`);
        } else {
            Logger.log(`تم حجز الميعاد ${slot.day} ${slot.timeSlotHeader} للطالب ${newStudentId}.`);
        }
    });

    Logger.log("اكتملت عملية حفظ الطالب الجديد بنجاح.");
    return { success: "تم تسجيل الطالب الجديد بنجاح." };

  } catch (e) {
    Logger.log("خطأ في saveData: " + e.message);
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
          subscriptionId: String(row[0] || '').trim(), // Subscription ID (A)
          packageName: String(row[2] || '').trim(),    // اسم الباقة (C)
          teacherId: String(row[3] || '').trim(),      // Teacher ID (D)
          startDate: row[4],                         // تاريخ بداية الاشتراك (E)
          endDate: row[5],                           // تاريخ نهاية الاشتراك المتوقع (F)
          attendedSessions: row[6],                  // عدد الحصص الحاضرة (G)
          renewalStatus: String(row[7] || '').trim(), // الحالة التفصيلية للتجديد (H)
          lastRenewalDate: row[8],                   // تاريخ آخر تجديد (I)
          totalSubscriptionAmount: row[9],           // مبلغ الاشتراك الكلي (J)
          paidAmount: row[10],                       // المبلغ المدفوع حتى الآن (K)
          remainingAmount: row[11]                   // المبلغ المتبقي (L)
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
// 4. الدوال الخاصة بصفحة تعديل الطلاب (Edit Student Page)
// ==============================================================================

/**
 * تجلب كافة البيانات الخاصة بطالب واحد (من الطلاب، الاشتراكات الحالية، والمواعيد).
 * هذه الدالة هي نسخة متخصصة من getAllStudentsData لطالب واحد.
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

  // 1. البحث عن الطالب في شيت "الطلاب" أولاً
  let studentRowIndex = -1;
  let studentBasicInfo = null;
  for (let i = 1; i < studentData.length; i++) {
    if (String(studentData[i][0] || '').trim() === String(studentId).trim()) { // العمود A: Student ID
      studentBasicInfo = {
        rowIndex: i + 1, // الصف في الشيت (1-based)
        studentID: String(studentData[i][0] || '').trim(),
        name: String(studentData[i][1] || '').trim(),
        age: studentData[i][2],
        phone: String(studentData[i][3] || '').trim(), // رقم الهاتف ولي الأمر
        studentPhone: String(studentData[i][4] || '').trim(), // رقم هاتف الطالب
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
 * تحديث بيانات الطالب في شيت "الطلاب" و "الاشتراكات الحالية"، وإعادة تخصيص المواعيد.
 *
 * @param {Object} updatedData - كائن يحتوي على البيانات المحدثة للطالب،
 * بالإضافة إلى oldTeacher, oldDay1, oldTime1, oldDay2, oldTime2 (لتحرير المواعيد القديمة).
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateStudentDataWithReassignment(updatedData) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");
    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' غير موجود.");
    if (!teachersSheet) throw new Error("شيت 'المعلمين' غير موجود.");

    const studentId = updatedData.studentID;
    const studentRowInStudentsSheet = updatedData.rowIndex; // صف الطالب في شيت "الطلاب"
    if (!studentId || !studentRowInStudentsSheet) {
      throw new Error("بيانات الطالب (Student ID أو رقم الصف) غير كاملة للتحديث.");
    }

    // 1. تحديث بيانات الطالب في شيت "الطلاب"
    // الأعمدة: Student ID(A), اسم الطالب(B), السن(C), رقم الهاتف(D), رقم هاتف الطالب(E), البلد(F), تاريخ التسجيل(G), الحالة الأساسية للطالب(H), ملاحظات(I)
    studentsSheet.getRange(studentRowInStudentsSheet, 2).setValue(updatedData.editName); // اسم الطالب (B)
    studentsSheet.getRange(studentRowInStudentsSheet, 3).setValue(updatedData.editAge);   // السن (C)
    // رقم الهاتف (D) لا يتم تحديثه من هنا لتجنب التعقيد في البحث بـ ID
    // يمكن إضافة تحديثات لأعمدة أخرى هنا إذا تم تعديلها في الواجهة

    // 2. تحديث بيانات الاشتراك في شيت "الاشتراكات الحالية"
    const currentSubscription = getStudentSubscriptionByStudentId(studentId); // دالة مساعدة جديدة
    if (!currentSubscription) {
        throw new Error(`لم يتم العثور على اشتراك للطالب ID ${studentId} لتحديثه.`);
    }

    const subscriptionRowInSheet = currentSubscription.rowIndex;
    const newTeacherId = getTeacherIdByName(updatedData.editTeacher); // جلب الـ ID من الاسم
    if (!newTeacherId) {
        throw new Error(`لم يتم العثور على Teacher ID للمعلم الجديد: ${updatedData.editTeacher}`);
    }

    // الأعمدة: Subscription ID(A), Student ID(B), اسم الباقة(C), Teacher ID(D), تاريخ بداية الاشتراك(E), تاريخ نهاية الاشتراك المتوقع(F), عدد الحصص الحاضرة(G), الحالة التفصيلية للتجديد(H), تاريخ آخر تجديد(I), مبلغ الاشتراك الكلي(J), المبلغ المدفوع حتى الآن(K), المبلغ المتبقي(L), ملاحظات خاصة بالاشتراك(M)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 3).setValue(updatedData.editSubscriptionPackage); // اسم الباقة (C)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 4).setValue(newTeacherId);                       // Teacher ID (D)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 8).setValue(updatedData.editPaymentStatus);      // الحالة التفصيلية للتجديد (H)

    // 3. معالجة إعادة تخصيص المواعيد في شيت "المواعيد المتاحة للمعلمين"
    // أ. تحرير المواعيد القديمة المحجوزة لهذا الطالب
    const oldBookedSlotsByStudent = getBookedSlotsByStudentId(teachersAvailableSlotsSheet, studentId); // دالة مساعدة جديدة
    oldBookedSlotsByStudent.forEach(slot => {
        const releaseResult = releaseTeacherSlot(
            teachersAvailableSlotsSheet,
            slot.teacherId, // المعلم الذي كان الميعاد محجوزاً لديه
            slot.day,
            slot.timeSlotHeader,
            studentId // تأكيد أن الطالب هو من حجز
        );
        if (releaseResult.error) {
            Logger.log(`خطأ في تحرير الميعاد القديم ${slot.day} ${slot.timeSlotHeader} للطالب ${studentId}: ${releaseResult.error}`);
            // يمكن الاستمرار أو رمي خطأ بناءً على مدى أهمية تحرير المواعيد القديمة
        }
    });

    // ب. حجز المواعيد الجديدة
    const newBookedSlotsToAssign = [];
    let bookingType = (updatedData.editPaymentStatus === "حلقة تجريبية") ? "تجريبي" : "عادي";

    if (updatedData.editDay1 && updatedData.editTime1) {
        newBookedSlotsToAssign.push({
            teacherId: newTeacherId, // المعلم الجديد
            day: updatedData.editDay1,
            timeSlotHeader: updatedData.editTime1,
            studentId: studentId,
            bookingType: bookingType
        });
    }
    if (updatedData.editDay2 && updatedData.editTime2) {
        newBookedSlotsToAssign.push({
            teacherId: newTeacherId, // المعلم الجديد
            day: updatedData.editDay2,
            timeSlotHeader: updatedData.editTime2,
            studentId: studentId,
            bookingType: bookingType
        });
    }

    newBookedSlotsToAssign.forEach(slot => {
        const bookResult = bookTeacherSlot( // إعادة استخدام دالة bookTeacherSlot
            teachersAvailableSlotsSheet,
            slot.teacherId,
            slot.day,
            slot.timeSlotHeader,
            slot.studentId,
            slot.bookingType
        );
        if (bookResult.error) {
            Logger.log(`خطأ في حجز الميعاد الجديد ${slot.day} ${slot.timeSlotHeader} للطالب ${studentId}: ${bookResult.error}`);
            throw new Error(`تعذر حجز الميعاد الجديد ${slot.day} ${slot.timeSlotHeader}: ${bookResult.error}`);
        }
    });

    Logger.log(`تم تحديث بيانات الطالب ${studentId} وإعادة تخصيص المواعيد بنجاح.`);
    return { success: "تم حفظ التعديلات بنجاح." };

  } catch (e) {
    Logger.log("خطأ في updateStudentDataWithReassignment: " + e.message);
    return { error: `فشل حفظ التعديلات: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * دالة مساعدة: تجلب تفاصيل اشتراك طالب واحد من شيت "الاشتراكات الحالية".
 * @param {string} studentId - Student ID للطالب.
 * @returns {Object|null} كائن الاشتراك أو null.
 */
function getStudentSubscriptionByStudentId(studentId) {
    const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
    if (!subscriptionsSheet) return null; // يجب أن يتم التحقق في الدوال الرئيسية

    const data = subscriptionsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][1] || '').trim() === String(studentId).trim()) { // العمود B: Student ID
            return {
                rowIndex: i + 1, // الصف في الشيت (1-based)
                subscriptionId: String(data[i][0] || '').trim(),
                packageName: String(data[i][2] || '').trim(),
                teacherId: String(data[i][3] || '').trim(),
                // ... يمكن إضافة باقي تفاصيل الاشتراك هنا إذا لزم الأمر
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
        targetCell.setValue(''); // جعل الخلية فارغة
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
  const subscriptionsSheet = SpreadsheetApp.GActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
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
    const reason = "حذف من النظام"; // يمكن تخصيص السبب أكثر
    
    // يمكن هنا جلب تفاصيل الاشتراك الأخيرة من شيت "الاشتراكات الحالية" لإضافتها للأرشيف
    // مثلاً: const lastSubscriptionDetails = getStudentSubscriptionByStudentId(studentIdToDelete);
    // وتضمينها في عمود "تفاصيل الاشتراك وقت الأرشفة"

    archiveSheet.appendRow([
      ...studentDataToArchive, // كل الأعمدة الأصلية من شيت الطلاب
      archivedDate,           // تاريخ الأرشفة
      reason,                 // سبب الأرشفة
      ""                      // تفاصيل الاشتراك وقت الأرشفة
    ]);
    Logger.log(`تم أرشفة الطالب ${studentDataToArchive[1]} (ID: ${studentIdToDelete}).`);


    // 4. حذف صف الطالب من شيت "الطلاب"
    studentsSheet.deleteRow(studentRowInStudentsSheet);
    Logger.log(`تم حذف صف الطالب ${studentIdToDelete} من شيت 'الطلاب'.`);

    // 5. حذف اشتراك الطالب من شيت "الاشتراكات الحالية"
    const studentSubscription = getStudentSubscriptionByStudentId(studentIdToDelete);
    if (studentSubscription) {
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

  // تخطي صفوف العناوين
  const students = studentData.slice(1);
  const subscriptions = subscriptionsData.slice(1);
  const teachers = teachersData.slice(1);

  const stats = {
    totalStudents: students.length,
    registeredStudents: 0,   // الطلاب المشتركين بـ "مشترك"
    trialStudents: 0,        // الطلاب بحالة "تجريبي"
    pendingStudents: 0,      // الطلاب بحالة "قيد التسجيل" أو "معلق"
    studentsByPackage: {},
    renewalStatus: {
      needsRenewal: 0,
      expired: 0,
      overLimit: 0,
      renewed: 0,
      trial: 0,              // عدد الطلاب في الحالة التجريبية
      notSubscribed: 0       // الطلاب بحالة "لم يشترك"
    },
    studentsByTeacher: {},   // key: Teacher Name, value: count of students
    recentlyRegistered: {
      last7Days: 0,
      last30Days: 0
    },
    renewalNextWeek: 0,
    renewalNextWeekStudents: [] // قائمة بالطلاب الذين يحتاجون للتجديد الأسبوع القادم
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0); // لتعيين الوقت إلى بداية اليوم لتجنب مشاكل فارق التوقيت في الحسابات
  
  const nextWeek = new Date(today);
  nextWeek.setDate(today.getDate() + 7);


  // بناء خريطة لربط Teacher ID باسم المعلم
  const teacherIdToNameMap = new Map();
  teachers.forEach(row => {
      const teacherId = String(row[0] || '').trim(); // العمود A: Teacher ID
      const teacherName = String(row[1] || '').trim(); // العمود B: اسم المعلم
      if (teacherId) {
          teacherIdToNameMap.set(teacherId, teacherName);
      }
  });


  // 1. حساب إحصائيات الطلاب من شيت "الطلاب"
  students.forEach(row => {
    const registrationDate = row[6]; // العمود G: تاريخ التسجيل
    const basicStatus = String(row[7] || '').trim(); // العمود H: الحالة الأساسية للطالب

    if (basicStatus === "مشترك") {
      stats.registeredStudents++;
    } else if (basicStatus === "تجريبي") {
      stats.trialStudents++;
    } else if (basicStatus === "قيد التسجيل" || basicStatus === "معلق") {
      stats.pendingStudents++;
    }

    // الطلاب المسجلون حديثًا (بناءً على تاريخ التسجيل)
    if (registrationDate instanceof Date) {
      const regDate = new Date(registrationDate);
      regDate.setHours(0, 0, 0, 0); 
      const diffTime = Math.abs(today.getTime() - regDate.getTime());
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays <= 7) {
        stats.recentlyRegistered.last7Days++;
      }
      if (diffDays <= 30) {
        stats.recentlyRegistered.last30Days++;
      }
    }
  });


  // 2. حساب إحصائيات من شيت "الاشتراكات الحالية"
  subscriptions.forEach(row => {
    const studentId = String(row[1] || '').trim(); // العمود B: Student ID
    const packageName = String(row[2] || '').trim(); // العمود C: اسم الباقة
    const teacherId = String(row[3] || '').trim(); // العمود D: Teacher ID
    const renewalStatus = String(row[7] || '').trim(); // العمود H: الحالة التفصيلية للتجديد
    const lastRenewalDate = row[8]; // العمود I: تاريخ آخر تجديد

    if (renewalStatus === "يحتاج للتجديد") {
      stats.renewalStatus.needsRenewal++;
    } else if (renewalStatus === "انتهت مدة الاشتراك") {
      stats.renewalStatus.expired++;
    } else if (renewalStatus === "تجاوز الحد") {
      stats.renewalStatus.overLimit++;
    } else if (renewalStatus === "تم التجديد") {
      stats.renewalStatus.renewed++;
    } else if (renewalStatus === "تجريبي") { // حالة تجريبي في الاشتراكات
        stats.renewalStatus.trial++;
    } else if (renewalStatus === "لم يشترك") { // حالة لم يشترك في الاشتراكات
        stats.renewalStatus.notSubscribed++;
    }


    // الطلاب حسب الباقة
    if (packageName) {
      stats.studentsByPackage[packageName] = (stats.studentsByPackage[packageName] || 0) + 1;
    }

    // الطلاب لكل معلم
    if (teacherId) {
        const teacherName = teacherIdToNameMap.get(teacherId) || teacherId; // استخدم الاسم
        if (!stats.studentsByTeacher[teacherName]) {
            stats.studentsByTeacher[teacherName] = new Set();
        }
        stats.studentsByTeacher[teacherName].add(studentId);
    }

    // حساب الطلاب الذين يحتاجون للتجديد الأسبوع القادم
    if (lastRenewalDate instanceof Date && renewalStatus === "تم التجديد") { // فقط الذين هم "تم التجديد" حالياً
      const renewalCheckDate = new Date(lastRenewalDate);
      renewalCheckDate.setMonth(renewalCheckDate.getMonth() + 1); // بعد شهر من آخر تجديد
      renewalCheckDate.setHours(0, 0, 0, 0);

      if (renewalCheckDate >= today && renewalCheckDate <= nextWeek) {
        stats.renewalNextWeek++;
        // جلب اسم الطالب ورقم الهاتف من شيت "الطلاب" (نحتاج أن نبحث عنه)
        const studentRow = students.find(s => String(s[0]).trim() === studentId); // العمود A هو Student ID
        if (studentRow) {
            stats.renewalNextWeekStudents.push({
                name: String(studentRow[1] || '').trim(), // اسم الطالب (B)
                phone: String(studentRow[3] || '').trim() // رقم الهاتف (D)
            });
        }
      }
    }
  });

  // تحويل Set إلى عدد صحيح لعدد الطلاب لكل معلم
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
function getAttendanceStudentsForTeacherAndDay(teacherId, day) { // تم تغيير الاسم ليكون أكثر دقة
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور"); // للتحقق من الحضور المسبق

  if (!studentsSheet) return { error: "شيت 'الطلاب' غير موجود." };
  if (!teachersAvailableSlotsSheet) return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود." };
  if (!attendanceLogSheet) { // لو مش موجود، ممكن ننشئه أو نبلغ بتحذير
      Logger.log("تحذير: شيت 'سجل الحضور' غير موجود. قد لا يتم التحقق من الحضور المسبق.");
  }


  // جلب بيانات الطلاب (ID -> Name)
  const studentIdToNameMap = new Map();
  const studentData = studentsSheet.getDataRange().getValues();
  studentData.forEach(row => {
      const id = String(row[0] || '').trim(); // العمود A: Student ID
      const name = String(row[1] || '').trim(); // العمود B: اسم الطالب
      if (id) studentIdToNameMap.set(id, name);
  });

  // جلب بيانات الحضور المسجلة لليوم الحالي لمنع تكرار التسجيل في الواجهة
  const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const markedAttendanceToday = new Set(); // Key: studentID_timeSlotHeader
  if (attendanceLogSheet) {
      const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
      attendanceLogData.forEach(row => {
          const logDate = row[3] ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : ''; // العمود D: تاريخ الحصة
          const logStudentID = String(row[1] || '').trim(); // العمود B: Student ID
          const logTimeSlot = String(row[2] || '').trim(); // العمود C: وقت الحصة

          if (logDate === todayFormatted) {
              markedAttendanceToday.add(`${logStudentID}_${logTimeSlot}`);
          }
      });
  }


  const studentsForAttendance = [];
  const teacherSlotsData = teachersAvailableSlotsSheet.getDataRange().getValues();
  const headers = teacherSlotsData[0]; // رؤوس أعمدة المواعيد في شيت المواعيد

  const timeSlotCols = [];
  const startColIndexForSlots = 2; // العمود C
  for (let i = startColIndexForSlots; i < headers.length; i++) {
    const header = String(headers[i] || '').trim();
    if (header) timeSlotCols.push({ index: i, header: header });
  }

  // البحث عن صفوف المعلم واليوم المطابقين
  for (let i = 1; i < teacherSlotsData.length; i++) {
    const row = teacherSlotsData[i];
    const currentTeacherId = String(row[0] || '').trim(); // العمود A: Teacher ID
    const currentDayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم

    if (currentTeacherId === teacherId && currentDayOfWeek === day) {
      timeSlotCols.forEach(colInfo => {
        const slotValue = String(row[colInfo.index] || '').trim();
        const timeSlotHeader = colInfo.header;

        // إذا كانت الخلية تحتوي على Student ID (محجوزة لطالب)
        if (slotValue.startsWith("STD") || slotValue.startsWith("p ")) {
            const studentIdInCell = slotValue;
            const studentName = studentIdToNameMap.get(studentIdInCell) || studentIdInCell; // جلب الاسم

            const isMarked = markedAttendanceToday.has(`${studentIdInCell}_${timeSlotHeader}`);

            studentsForAttendance.push({
                studentID: studentIdInCell,
                name: studentName,
                timeSlot: timeSlotHeader,
                isMarked: isMarked // لتحديد ما إذا كان تم تسجيل حضوره بالفعل
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
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب"); // لجلب اسم الطالب
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  let attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الحضور");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");

    // التحقق من وجود شيت سجل الحضور، وإلا إنشاءه
    if (!attendanceLogSheet) {
      attendanceLogSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("سجل الحضور");
      attendanceLogSheet.appendRow([
        "Attendance ID", "Student ID", "Teacher ID", "Subscription ID",
        "تاريخ الحصة", "وقت الحصة", "اليوم", "حالة الحضور", "نوع الحصة", "ملاحظات المعلم"
      ]);
    }

    const today = new Date();
    const todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // التحقق من عدم تسجيل الحضور مسبقًا لنفس الطالب في نفس اليوم والوقت في نفس التاريخ
    const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
    for (let i = 1; i < attendanceLogData.length; i++) {
      const logRow = attendanceLogData[i];
      const logStudentID = String(logRow[1] || '').trim(); // العمود B: Student ID
      const logTeacherID = String(logRow[2] || '').trim(); // العمود C: Teacher ID
      const logDate = logRow[4] ? Utilities.formatDate(logRow[4], Session.getScriptTimeZone(), "yyyy-MM-dd") : ''; // العمود E: تاريخ الحصة
      const logTimeSlot = String(logRow[5] || '').trim(); // العمود F: وقت الحصة

      if (logStudentID === studentId && logTeacherID === teacherId && logDate === todayFormatted && logTimeSlot === timeSlot) {
        return { error: "تم تسجيل الحضور لهذا الطالب في هذا الميعاد وهذا اليوم مسبقًا." };
      }
    }

    // جلب اسم الطالب (للتسجيل في السجل)
    const studentName = getStudentNameById(studentId); // دالة مساعدة
    if (!studentName) throw new Error(`لم يتم العثور على اسم الطالب بمعرف ${studentId}.`);

    // جلب Subscription ID ونوع الباقة (للتحديث وتحديد نوع الحصة)
    let subscriptionId = '';
    let packageName = '';
    let renewalStatus = '';
    let totalPackageSessions = 0; // إجمالي الحصص في الباقة
    
    const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
    let subscriptionRowIndex = -1;

    for (let i = 1; i < subscriptionsData.length; i++) {
        if (String(subscriptionsData[i][1] || '').trim() === studentId) { // العمود B: Student ID
            subscriptionRowIndex = i;
            subscriptionId = String(subscriptionsData[i][0] || '').trim(); // العمود A: Subscription ID
            packageName = String(subscriptionsData[i][2] || '').trim(); // العمود C: اسم الباقة
            renewalStatus = String(subscriptionsData[i][7] || '').trim(); // العمود H: الحالة التفصيلية للتجديد
            totalPackageSessions = getTotalSessionsForPackage(packageName); // دالة مساعدة
            break;
        }
    }
    if (subscriptionRowIndex === -1) throw new Error(`لم يتم العثور على اشتراك للطالب ID ${studentId}.`);


    // تحديد نوع الحصة (عادية / تجريبية)
    let classType = "عادية";
    if (renewalStatus === "تجريبي") {
        classType = "تجريبية";
    }

    // تسجيل الحضور في شيت "سجل الحضور"
    const attendanceId = generateUniqueAttendanceId(attendanceLogSheet); // دالة مساعدة
    // الأعمدة: Attendance ID, Student ID, Teacher ID, Subscription ID, تاريخ الحصة, وقت الحصة, اليوم, حالة الحضور, نوع الحصة, ملاحظات المعلم
    attendanceLogSheet.appendRow([
      attendanceId,             // A
      studentId,                // B
      teacherId,                // C
      subscriptionId,           // D
      today,                    // E
      timeSlot,                 // F
      day,                      // G
      "حضر",                    // H (حالة الحضور)
      classType,                // I (نوع الحصة)
      ""                        // J (ملاحظات المعلم)
    ]);
    Logger.log(`تم تسجيل الحضور لـ ${studentName} (ID: ${studentId}) في ${day} ${timeSlot}.`);


    // تحديث عدد الحصص الحاضرة في شيت "الاشتراكات الحالية" (فقط للحصص العادية)
    if (classType === "عادية") {
        const currentAttendedSessionsCell = subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7); // العمود G: عدد الحصص الحاضرة
        let currentSessions = currentAttendedSessionsCell.getValue();
        currentSessions = (typeof currentSessions === 'number') ? currentSessions : 0;
        subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7).setValue(currentSessions + 1);
        Logger.log(`تم تحديث عدد الحصص الحاضرة للطالب ${studentId} إلى ${currentSessions + 1}.`);

        // تحديث "الحالة التفصيلية للتجديد" إذا وصل الطالب إلى الحد الأقصى من الحصص
        if (totalPackageSessions > 0 && (currentSessions + 1) >= totalPackageSessions) {
            subscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue("يحتاج للتجديد"); // العمود H
            Logger.log(`حالة تجديد الطالب ${studentId} تم تحديثها إلى "يحتاج للتجديد".`);
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
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' غير موجود.");
    if (!teachersAvailableSlotsSheet) throw new Error("شيت 'المواعيد المتاحة للمعلمين' غير موجود.");
    if (!teachersSheet) throw new Error("شيت 'المعلمين' غير موجود.");

    const studentId = updatedData.studentID;
    const studentRowInStudentsSheet = updatedData.rowIndex; // صف الطالب في شيت "الطلاب"
    if (!studentId || !studentRowInStudentsSheet) {
      throw new Error("بيانات الطالب (Student ID أو رقم الصف) غير كاملة للتحديث.");
    }

    // 1. تحديث بيانات الطالب في شيت "الطلاب"
    // الأعمدة: Student ID(A), اسم الطالب(B), السن(C), رقم الهاتف(D), رقم هاتف الطالب(E), البلد(F), تاريخ التسجيل(G), الحالة الأساسية للطالب(H), ملاحظات(I)
    studentsSheet.getRange(studentRowInStudentsSheet, 2).setValue(updatedData.editName); // اسم الطالب (B)
    studentsSheet.getRange(studentRowInStudentsSheet, 3).setValue(updatedData.editAge);   // السن (C)
    studentsSheet.getRange(studentRowInStudentsSheet, 4).setValue(updatedData.editPhone); // رقم الهاتف (D) - تم إضافته
    // يمكن إضافة تحديثات لأعمدة أخرى هنا إذا تم تعديلها في الواجهة

    // 2. تحديث بيانات الاشتراك في شيت "الاشتراكات الحالية"
    const currentSubscription = getStudentSubscriptionByStudentId(studentId); // دالة مساعدة
    if (!currentSubscription) {
        throw new Error(`لم يتم العثور على اشتراك للطالب ID ${studentId} لتحديثه.`);
    }

    const subscriptionRowInSheet = currentSubscription.rowIndex;
    const newTeacherId = getTeacherIdByName(updatedData.editTeacher); // جلب الـ ID من الاسم
    if (!newTeacherId) {
        throw new Error(`لم يتم العثور على Teacher ID للمعلم الجديد: ${updatedData.editTeacher}`);
    }

    // الأعمدة: Subscription ID(A), Student ID(B), اسم الباقة(C), Teacher ID(D), تاريخ بداية الاشتراك(E), تاريخ نهاية الاشتراك المتوقع(F), عدد الحصص الحاضرة(G), الحالة التفصيلية للتجديد(H), تاريخ آخر تجديد(I), مبلغ الاشتراك الكلي(J), المبلغ المدفوع حتى الآن(K), المبلغ المتبقي(L), ملاحظات خاصة بالاشتراك(M)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 3).setValue(updatedData.editSubscriptionPackage); // اسم الباقة (C)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 4).setValue(newTeacherId);                       // Teacher ID (D)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 8).setValue(updatedData.editPaymentStatus);      // الحالة التفصيلية للتجديد (H)

    // 3. معالجة إعادة تخصيص المواعيد في شيت "المواعيد المتاحة للمعلمين"
    // أ. تحرير المواعيد القديمة المحجوزة بواسطة هذا الطالب
    const oldBookedSlotsByStudent = getBookedSlotsByStudentId(teachersAvailableSlotsSheet, studentId); // دالة مساعدة
    oldBookedSlotsByStudent.forEach(slot => {
        const releaseResult = releaseTeacherSlot(
            teachersAvailableSlotsSheet,
            slot.teacherId, // المعلم الذي كان الميعاد محجوزاً لديه
            slot.day,
            slot.timeSlotHeader,
            studentId // تأكيد أن الطالب هو من حجز
        );
        if (releaseResult.error) {
            Logger.log(`خطأ في تحرير الميعاد القديم ${slot.day} ${slot.timeSlotHeader} للطالب ${studentId}: ${releaseResult.error}`);
            // يمكن الاستمرار أو رمي خطأ بناءً على مدى أهمية تحرير المواعيد القديمة
        }
    });

    // ب. حجز المواعيد الجديدة
    const newBookedSlotsToAssign = [];
    let bookingType = (updatedData.editPaymentStatus === "تجريبي") ? "تجريبي" : "عادي"; // نوع الحجز بناءً على حالة التجديد

    if (updatedData.editDay1 && updatedData.editTime1) {
        newBookedSlotsToAssign.push({
            teacherId: newTeacherId, // المعلم الجديد
            day: updatedData.editDay1,
            timeSlotHeader: updatedData.editTime1,
            studentId: studentId,
            bookingType: bookingType
        });
    }
    if (updatedData.editDay2 && updatedData.editTime2) {
        newBookedSlotsToAssign.push({
            teacherId: newTeacherId, // المعلم الجديد
            day: updatedData.editDay2,
            timeSlotHeader: updatedData.editTime2,
            studentId: studentId,
            bookingType: bookingType
        });
    }

    newBookedSlotsToAssign.forEach(slot => {
        const bookResult = bookTeacherSlot( // إعادة استخدام دالة bookTeacherSlot
            teachersAvailableSlotsSheet,
            slot.teacherId,
            slot.day,
            slot.timeSlotHeader,
            slot.studentId,
            slot.bookingType
        );
        if (bookResult.error) {
            Logger.log(`خطأ في حجز الميعاد الجديد ${slot.day} ${slot.timeSlotHeader} للطالب ${studentId}: ${bookResult.error}`);
            throw new Error(`تعذر حجز الميعاد الجديد ${slot.day} ${slot.timeSlotHeader}: ${bookResult.error}`);
        }
    });

    Logger.log(`تم تحديث بيانات الطالب ${studentId} وإعادة تخصيص المواعيد بنجاح.`);
    return { success: "تم حفظ التعديلات بنجاح." };

  } catch (e) {
    Logger.log("خطأ في updateStudentDataWithReassignment: " + e.message);
    return { error: `فشل حفظ التعديلات: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
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
                packageName: String(data[i][2] || '').trim(),    // العمود C
                teacherId: String(data[i][3] || '').trim(),      // العمود D
                // ... يمكن إضافة باقي تفاصيل الاشتراك هنا إذا لزم الأمر
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
        targetCell.setValue(''); // جعل الخلية فارغة
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
