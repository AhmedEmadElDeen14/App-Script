/**
 * ملف code.gs
 * هذا الملف يحتوي على دوال App Script لنظام إدارة أكاديمية رفاق.
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
      .setTitle('نظام إدارة أكاديمية رفاق'); // عنوان يظهر في تبويبة المتصفح
}

/**
 * دالة مساعدة لتمكين الواجهة الأمامية من الحصول على رابط تطبيق الويب الحالي.
 * @returns {string} رابط تطبيق الويب.
 */
function getBaseUrl() {
  return ScriptApp.getService().getUrl();
}


/**
 * دالة مساعدة لتحويل تنسيق المواعيد (9ص -> 09:00، p 3 -> 15:00) للفرز والتعامل في Code.gs.
 * تفترض أن "9ص" أو "p 3" تمثل بداية الميعاد.
 * @param {string} timeString - سلسلة الوقت (مثل "9ص", "1:30 م", "p 3", "09:00 - 09:30").
 * @returns {string} الوقت بتنسيق 24 ساعة (HH:mm).
 */
function convertTo24HourFormat(timeString) {
    if (typeof timeString !== 'string' || timeString.trim() === '') return '00:00';
    timeString = timeString.trim();

    const parts = timeString.split(' - ');
    if (parts.length > 1 && parts[0].includes(':')) {
        timeString = parts[0];
    }

    if (timeString.toLowerCase().startsWith('p ')) {
        const hourPart = parseInt(timeString.substring(2));
        if (!isNaN(hourPart) && hourPart >= 1 && hourPart <= 12) {
            const hour24 = (hourPart === 12) ? 12 : hourPart + 12;
            return `${hour24.toString().padStart(2, '0')}:00`;
        }
    } else if (timeString.toLowerCase().endsWith('ص')) {
        const hourPart = parseInt(timeString.replace('ص', ''));
        if (!isNaN(hourPart) && hourPart >= 1 && hourPart <= 12) {
            const hour24 = (hourPart === 12) ? 0 : hourPart;
            return `${hour24.toString().padStart(2, '0')}:00`;
        }
    } else if (timeString.toLowerCase().endsWith('م')) {
        const hourPart = parseInt(timeString.replace('م', ''));
        if (!isNaN(hourPart) && hourPart >= 1 && hourPart <= 12) {
            const hour24 = (hourPart === 12) ? 12 : hourPart + 12;
            return `${hour24.toString().padStart(2, '0')}:00`;
        }
    } else if (timeString.includes(':')) {
        const [hours, minutes] = timeString.split(':').map(Number);
        if (!isNaN(hours) && !isNaN(minutes)) {
            return `${hours.toString().padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
        }
    } else {
        const hourPart = parseInt(timeString);
        if (!isNaN(hourPart) && hourPart >= 0 && hourPart <= 23) {
            return `${hourPart.toString().padStart(2, '0')}:00`;
        }
    }
    Logger.log("Warning: convertTo24HourFormat (Code.gs) could not parse: " + timeString);
    return '00:00'; // Fallback
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
 * @returns {Object|null} كائن يحتوي على تفاصيل الباقة (السعر، عدد الحصص الكلي، نوع الباقة، مدة الحصة (دقيقة)، عدد الحصص الأسبوعية) أو null.
 */
function getPackageDetails(packageName) {
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");
  if (!packagesSheet) {
    Logger.log("خطأ: لم يتم العثور على شيت 'الباقات' في getPackageDetails.");
    return null;
  }
  const data = packagesSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // العمود A: اسم الباقة، العمود B: مدة الحصة، العمود C: عدد الحصص الكلي، العمود D: السعر، العمود E: نوع الباقة، العمود F: عدد الحصص الأسبوعية
    if (String(data[i][0] || '').trim() === String(packageName).trim()) {
      return {
        'اسم الباقة': String(data[i][0] || '').trim(),
        'مدة الحصة (دقيقة)': data[i][1], // العمود B (مؤشر 1)
        'عدد الحصص الكلي': data[i][2], // العمود C (مؤشر 2)
        'السعر': data[i][3], // العمود D (مؤشر 3)
        'نوع الباقة': String(data[i][4] || '').trim(), // العمود E (مؤشر 4)
        'عدد الحصص الأسبوعية': data[i][5] || 0 // العمود F (مؤشر 5)
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
    return { error: `الميعاد ${day} ${timeSlotHeader} محجوز بالفعل لطالب آخر: ${currentSlotValue}.` };
  } else { // لو الخلية فارغة (غير متاحة)
    return { error: `الميعاد ${day} ${timeSlotHeader} غير متاح للحجز.` };
  }

  // هذه التعليقات كانت مكررة من قبل، وظيفتها كانت في الإصدارات القديمة من الدالة
  // const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  // sheet.getRange(teacherRowIndex + 1, 4).setValue(today); // تاريخ الحجز (D)
  // sheet.getRange(teacherRowIndex + 1, 5).setValue(bookingType); // نوع الحجز (E)

  return { success: `تم حجز الميعاد بنجاح للطالب ${studentId}.` };
}



/**
 * دالة مساعدة: تحسب رأس عمود الميعاد التالي مباشرة (بفارق 30 دقيقة) بناءً على الميعاد الحالي.
 * تفترض أن المواعيد في الشيت متسلسلة بفروقات 30 دقيقة.
 *
 * @param {string} currentTimeSlotHeader - رأس عمود الميعاد الحالي (مثلاً: "09:00 - 09:30").
 * @returns {string|null} رأس عمود الميعاد التالي (مثلاً: "09:30 - 10:00") أو null إذا لم يكن هناك موعد تالي صالح.
 */
function getNextTimeSlotHeader(currentTimeSlotHeader) {
  if (typeof currentTimeSlotHeader !== 'string' || currentTimeSlotHeader.trim() === '') {
    Logger.log("Invalid currentTimeSlotHeader provided: " + currentTimeSlotHeader);
    return null;
  }

  const parts = currentTimeSlotHeader.split(' - ');
  // يجب أن يكون رأس الميعاد بالتنسيق "HH:MM - HH:MM"
  if (parts.length !== 2 || !parts[0].includes(':') || !parts[1].includes(':')) {
    Logger.log("Invalid time slot header format: " + currentTimeSlotHeader);
    return null;
  }

  // نأخذ وقت نهاية الميعاد الحالي كبداية للميعاد التالي
  const [endHourStr, endMinuteStr] = parts[1].split(':');
  const currentEndHour = parseInt(endHourStr);
  const currentEndMinute = parseInt(endMinuteStr);

  if (isNaN(currentEndHour) || isNaN(currentEndMinute)) {
    Logger.log("Could not parse end time from header: " + currentTimeSlotHeader);
    return null;
  }

  // حساب بداية الميعاد التالي (نهاية الميعاد الحالي)
  const nextSlotStartMinutes = (currentEndHour * 60) + currentEndMinute;
  const nextSlotStartHour = Math.floor(nextSlotStartMinutes / 60);
  const nextSlotStartMinute = nextSlotStartMinutes % 60;

  // حساب نهاية الميعاد التالي (بعد 30 دقيقة من بدايته)
  const nextSlotEndMinutes = nextSlotStartMinutes + 30;
  const nextSlotEndHour = Math.floor(nextSlotEndMinutes / 60);
  const nextSlotEndMinute = nextSlotEndMinutes % 60;

  // تنسيق الموعد التالي
  const nextStartTimeFormatted = `${String(nextSlotStartHour).padStart(2, '0')}:${String(nextSlotStartMinute).padStart(2, '0')}`;
  const nextEndTimeFormatted = `${String(nextSlotEndHour).padStart(2, '0')}:${String(nextSlotEndMinute).padStart(2, '0')}`;

  // تأكد أن الموعد التالي لا يتجاوز نهاية اليوم (23:30 - 00:00 هو أقصى ميعاد ممكن)
  // إذا تخطى 23:30، أو وصل لـ 24:00 (منتصف الليل)
  if (nextSlotStartHour > 23 || (nextSlotStartHour === 23 && nextSlotStartMinute > 30)) {
      Logger.log("Next time slot is beyond valid range (after 23:30): " + currentTimeSlotHeader);
      return null;
  }
  
  return `${nextStartTimeFormatted} - ${nextEndTimeFormatted}`;
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
    "لم يشترك",
    "مجاني" // <-- تم إضافة هذا الخيار
  ];
  Logger.log("getPaymentStatusList: Returned statuses: " + JSON.stringify(statuses));
  return statuses;
}


/**
 * حفظ بيانات الطالب الجديد.
 * تقوم بحفظ البيانات في شيت "الطلاب"، وإنشاء اشتراك في شيت "الاشتراكات الحالية"،
 * وتحديث المواعيد المحجوزة في شيت "المواعيد المتاحة للمعلمين".
 *
 * @param {Object} formData - كائن يحتوي على بيانات النموذج المرسلة من الواجهة الأمامية.
 * المتوقع من الواجهة الأمامية (من حقول التسجيل):
 * { regName, regAge, regPhone, regTeacher (اسم المعلم), regPaymentStatus (حالة الدفع),
 * regSubscriptionPackage (اسم الباقة), regSubscriptionType (نوع الاشتراك),
 * regSlots: Array<{day: string, time: string}> (مصفوفة المواعيد) }
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function saveData(formData) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // التحقق من وجود الشيتات الضرورية
    if (!studentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب'.");
    if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    if (!teachersSheet) throw new Error("لم يتم العثور على شيت 'المعلمين'.");
    if (!packagesSheet) throw new Error("لم يتم العثور على شيت 'الباقات'.");

    // 1. جلب Teacher ID من اسم المعلم
    const teacherId = getTeacherIdByName(formData.regTeacher);
    if (!teacherId) throw new Error(`لم يتم العثور على Teacher ID للمعلم: ${formData.regTeacher}`);

    // 2. حفظ الطالب في شيت "الطلاب"
    const newStudentId = generateUniqueStudentId(studentsSheet);
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    let studentBasicStatus = "قيد التسجيل";
    if (formData.regPaymentStatus === "تم الدفع") {
        studentBasicStatus = "مشترك";
    } else if (formData.regPaymentStatus === "مجاني") { // <-- جديد
        studentBasicStatus = "مشترك"; // الطالب المجاني يعتبر مشترك
    } else if (formData.regPaymentStatus === "حلقة تجريبية") {
        studentBasicStatus = "تجريبي";
    } else if (formData.regPaymentStatus === "لم يشترك" || formData.regPaymentStatus === "تم دفع جزء") {
        studentBasicStatus = "معلق";
    }

    studentsSheet.appendRow([
      newStudentId,                       // Student ID (A)
      formData.regName,                   // اسم الطالب (B)
      formData.regAge,                    // السن (C)
      String(formData.regPhone).trim(),   // رقم الهاتف (ولي الأمر) (D)
      "",                                 // رقم هاتف الطالب (إن وجد) (E)
      "",                                 // البلد (F)
      today,                              // تاريخ التسجيل (G)
      studentBasicStatus,                 // الحالة الأساسية للطالب (H)
      ""                                  // ملاحظات (I)
    ]);
    Logger.log(`تم حفظ الطالب ${formData.regName} (ID: ${newStudentId}) في شيت 'الطلاب'.`);

    // 3. إنشاء اشتراك في شيت "الاشتراكات الحالية"
    const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet);
    const packageName = formData.regSubscriptionPackage;
    const subscriptionType = formData.regSubscriptionType;
    const packageDetails = getPackageDetails(packageName);

    let subscriptionRenewalStatus = "لم يشترك";
    let totalClassesAttended = 0;
    let subscriptionPackageType = "";
    let subscriptionAmount = 0;
    let paidAmount = 0;
    let remainingAmount = 0;
    let sessionDuration = 0; // مدة الحصة بالدقائق

    if (packageDetails) {
        subscriptionAmount = packageDetails['السعر'] || 0;
        subscriptionPackageType = packageDetails['نوع الباقة'] || "";
        sessionDuration = packageDetails['مدة الحصة (دقيقة)'] || 0;

        if (formData.regPaymentStatus === "تم الدفع") {
            subscriptionRenewalStatus = "تم التجديد";
            paidAmount = subscriptionAmount;
            remainingAmount = 0;
        } else if (formData.regPaymentStatus === "حلقة تجريبية") {
            subscriptionRenewalStatus = "تجريبي";
            paidAmount = 0;
            remainingAmount = 0;
        } else if (formData.regPaymentStatus === "تم دفع جزء") {
            subscriptionRenewalStatus = "تم دفع جزء";
            paidAmount = 0;
            remainingAmount = subscriptionAmount;
        } else if (formData.regPaymentStatus === "لم يتم الدفع") {
            subscriptionRenewalStatus = "لم يتم الدفع";
            paidAmount = 0;
            remainingAmount = subscriptionAmount;
        }
    } else {
        if (formData.regPaymentStatus === "حلقة تجريبية") {
            subscriptionRenewalStatus = "تجريبي";
        } else {
            subscriptionRenewalStatus = "لم يشترك";
        }
        Logger.log(`تحذير: باقة '${packageName}' غير موجودة. الاشتراك سيُسجل بقيم افتراضية للمبلغ.`);
    }

    // حساب تاريخ نهاية الاشتراك المتوقع
    let endDate = "";
    if (packageDetails) {
      if (packageDetails['نوع الباقة'] === "شهري") {
          const startDate = new Date(today);
          startDate.setMonth(startDate.getMonth() + 1);
          endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (packageDetails['نوع الباقة'] === "نصف سنوي") {
          const startDate = new Date(today);
          startDate.setMonth(startDate.getMonth() + 6);
          endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (packageDetails['نوع الباقة'] === "سنوي") {
          const startDate = new Date(today);
          startDate.setFullYear(startDate.getFullYear() + 1);
          endDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
    }

    // ترتيب الأعمدة في شيت "الاشتراكات الحالية": Subscription ID, Student ID, اسم الباقة, Teacher ID, تاريخ بداية الاشتراك, تاريخ نهاية الاشتراك المتوقع, عدد الحصص الحاضرة, الحالة التفصيلية للتجديد, تاريخ آخر تجديد, مبلغ الاشتراك الكلي, المبلغ المدفوع حتى الآن, المبلغ المتبقي, نوع الاشتراك, ملاحظات خاصة بالاشتراك
    subscriptionsSheet.appendRow([
      newSubscriptionId,           // A
      newStudentId,                // B
      packageName,                 // C
      teacherId,                   // D
      today,                       // E
      endDate,                     // F
      totalClassesAttended,        // G (عدد الحصص الحاضرة، يبدأ من صفر)
      subscriptionRenewalStatus,   // H
      today,                       // I (تاريخ آخر تجديد)
      subscriptionAmount,          // J
      paidAmount,                  // K
      remainingAmount,             // L
      subscriptionType,            // M
      ""                           // N (ملاحظات)
    ]);
    Logger.log(`تم إنشاء اشتراك (${newSubscriptionId}) للطالب ${newStudentId} في شيت 'الاشتراكات الحالية'.`);


    // 4. حجز المواعيد في شيت "المواعيد المتاحة للمعلمين"
    let bookingTypeForSlot = (formData.regPaymentStatus === "حلقة تجريبية") ? "تجريبي" : "عادي";

    if (!formData.regSlots || formData.regSlots.length === 0) {
        Logger.log("لم يتم تحديد مواعيد لحجزها.");
    } else {
        // حجز كل المواعيد المحددة في regSlots
        for (let i = 0; i < formData.regSlots.length; i++) {
            const slot = formData.regSlots[i];
            if (slot.day && slot.time) {
                const initialTimeSlotHeader = slot.time;   
                if (initialTimeSlotHeader) {
                    // **أولاً: حجز الموعد الأساسي الذي تم اختياره من قبل المستخدم**
                    let result = bookTeacherSlot(
                        teachersAvailableSlotsSheet,
                        teacherId,
                        slot.day,
                        initialTimeSlotHeader, // نستخدم الميعاد الذي اختاره المستخدم مباشرة
                        newStudentId,
                        bookingTypeForSlot
                    );
                    if (result.error) {
                        Logger.log(`خطأ في حجز الميعاد الأول ${slot.day} ${initialTimeSlotHeader} للطالب ${newStudentId}: ${result.error}`);
                        throw new Error(`تعذر حجز الميعاد الأول ${slot.day} ${initialTimeSlotHeader}: ${result.error}`);
                    } else {
                        Logger.log(`تم حجز الميعاد الأول ${slot.day} ${initialTimeSlotHeader} للطالب ${newStudentId}.`);
                    }

                    // **ثانياً: إذا كانت مدة الحصة ساعة (60 دقيقة)، احجز الموعد التالي له مباشرةً**
                    if (sessionDuration === 60) {
                        const nextTimeSlotHeader = getNextTimeSlotHeader(initialTimeSlotHeader); // نحسب الميعاد التالي من الميعاد الأساسي المحجوز
                        if (nextTimeSlotHeader) {
                            result = bookTeacherSlot( // إعادة استخدام result
                                teachersAvailableSlotsSheet,
                                teacherId,
                                slot.day,
                                nextTimeSlotHeader,
                                newStudentId,
                                bookingTypeForSlot
                            );
                            if (result.error) {
                                Logger.log(`خطأ في حجز الميعاد الثاني ${slot.day} ${nextTimeSlotHeader} للطالب ${newStudentId}: ${result.error}`);
                                throw new Error(`تعذر حجز الميعاد الثاني ${slot.day} ${nextTimeSlotHeader}: ${result.error}`);
                            } else {
                                Logger.log(`تم حجز الميعاد الثاني ${slot.day} ${nextTimeSlotHeader} للطالب ${newStudentId}.`);
                            }
                        } else {
                            Logger.log(`تحذير: لم يتم العثور على ميعاد تالي صالح للميعاد الأول ${slot.day} ${initialTimeSlotHeader} لحصة مدتها ساعة.`);
                        }
                    }
                } else {
                    Logger.log(`تحذير: تنسيق وقت غير صالح للميعاد: '${slot.time}'. تم تخطي حجز هذا الميعاد.`);
                }
            }
        }
    }

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
 * أو {Object} كائن خطأ إذا كان شيت "الطلاب" غير موجود أو أي شيت أساسي آخر.
 */
function getAllStudentsData() {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");

  // التحقق من وجود الشيتات وإرجاع خطأ صريح إذا كانت مفقودة
  if (!studentsSheet) {
    Logger.log("خطأ: شيت 'الطلاب' غير موجود لـ getAllStudentsData.");
    return { error: "شيت 'الطلاب' غير موجود." };
  }
  if (!subscriptionsSheet) {
    Logger.log("خطأ: شيت 'الاشتراكات الحالية' غير موجود لـ getAllStudentsData.");
    return { error: "شيت 'الاشتراكات الحالية' غير موجود." };
  }
  if (!teachersSheet) {
    Logger.log("خطأ: شيت 'المعلمين' غير موجود لـ getAllStudentsData.");
    return { error: "شيت 'المعلمين' غير موجود." };
  }
  if (!teachersAvailableSlotsSheet) {
    Logger.log("خطأ: شيت 'المواعيد المتاحة للمعلمين' غير موجود لـ getAllStudentsData.");
    return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود." };
  }

  try {
      const studentData = studentsSheet.getDataRange().getValues();
      // إذا كان شيت الطلاب فارغاً تماماً (حتى لا يوجد رؤوس أعمدة)، فإن getDataRange().getValues() قد ترجع مصفوفة فارغة []
      if (studentData.length < 1) { // NEW CHECK: if sheet is completely empty
          Logger.log("شيت 'الطلاب' فارغ تمامًا أو لا يحتوي على بيانات.");
          return []; // إرجاع مصفوفة فارغة من الطلاب
      }
      // تخطي صفوف العناوين (يجب أن يكون الصف الأول هو رؤوس الأعمدة)
      const actualStudentData = studentData.slice(1);


      const allStudents = [];

      // جلب بيانات الاشتراكات لربطها بـ Student ID
      const subscriptionsMap = new Map(); // key: Student ID, value: { subscriptionDetails }
      const subscriptionsValues = subscriptionsSheet.getDataRange().getValues();
      if (subscriptionsValues.length > 1) { // NEW CHECK: ensure there's more than just header row
          subscriptionsValues.forEach((row, index) => {
            if (index === 0) return; // تخطي صف العناوين في الاشتراكات
            const studentID = String(row[1] || '').trim(); // العمود B: Student ID في الاشتراكات
            if (studentID) {
              subscriptionsMap.set(studentID, {
                subscriptionId: String(row[0] || '').trim(), // Subscription ID (A)
                packageName: String(row[2] || '').trim(),    // اسم الباقة (C)
                teacherId: String(row[3] || '').trim(),      // Teacher ID (D)
                startDate: row[4],                           // تاريخ بداية الاشتراك (E)
                endDate: row[5],                             // تاريخ نهاية الاشتراك المتوقع (F)
                attendedSessions: row[6],                    // عدد الحصص الحاضرة (G)
                renewalStatus: String(row[7] || '').trim(),  // الحالة التفصيلية للتجديد (H)
                lastRenewalDate: row[8],                     // تاريخ آخر تجديد (I)
                totalSubscriptionAmount: row[9],             // مبلغ الاشتراك الكلي (J)
                paidAmount: row[10],                         // المبلغ المدفوع حتى الآن (K)
                remainingAmount: row[11],                    // المبلغ المتبقي (L)
                subscriptionType: String(row[12] || '').trim() // نوع الاشتراك (M)
              });
            }
          });
      }


      // جلب بيانات المعلمين (للحصول على اسم المعلم من Teacher ID)
      const teacherIdToNameMap = new Map();
      const teachersValues = teachersSheet.getDataRange().getValues();
      if (teachersValues.length > 1) { // NEW CHECK: ensure there's more than just header row
          teachersValues.forEach(row => {
              const teacherId = String(row[0] || '').trim(); // العمود A: Teacher ID
              const teacherName = String(row[1] || '').trim(); // العمود B: اسم المعلم
              if (teacherId) {
                  teacherIdToNameMap.set(teacherId, teacherName);
              }
          });
      }


      // جلب بيانات المواعيد المحجوزة لكل طالب
      const studentBookedSlotsMap = new Map(); // key: Student ID, value: [{day, timeSlotHeader, teacherId}, ...]
      const availableSlotsValues = teachersAvailableSlotsSheet.getDataRange().getValues();
      if (availableSlotsValues.length > 1) { // NEW CHECK: ensure there's more than just header row
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
              const teacherIdInSlot = String(row[0] || '').trim(); // العمود A: Teacher ID
              const dayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم

              timeSlotHeaders.forEach(colInfo => {
                  const slotValue = String(row[colInfo.index] || '').trim();
                  const timeSlotHeader = colInfo.header;

                  // لو الخلية فيها Student ID (محجوز)
                  if (slotValue.startsWith("STD") || slotValue.startsWith("TRL") || slotValue.startsWith("p ")) {
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
      actualStudentData.forEach((row, index) => { // نمر على البيانات بعد تخطي صفوف العناوين
        const studentID = String(row[0] || '').trim();

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
            studentInfo.subscriptionType = subscriptionDetails.subscriptionType;
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
            studentInfo.subscriptionType = '';
        }

        // حساب الحصص المتبقية (إذا كانت الباقة محددة العدد)
        const totalSessions = getTotalSessionsForPackage(studentInfo.packageName); // دالة مساعدة
        let remainingSessions = 'غير متاح';
        if (totalSessions > 0 && typeof studentInfo.attendedSessions === 'number') {
            remainingSessions = totalSessions - studentInfo.attendedSessions;
        } else if (studentInfo.renewalStatus === "تجريبي") {
            remainingSessions = "تجريبي";
        } else if (studentInfo.packageName.includes("نصف سنوي") || studentInfo.packageName.includes("سنوي") || studentInfo.packageName.includes("مخصص")) {
            remainingSessions = "غير محدد";
        }
        studentInfo.remainingSessions = remainingSessions;


        // دمج جميع المواعيد المحجوزة
        const bookedSlots = studentBookedSlotsMap.get(studentID) || [];
        studentInfo.allBookedScheduleSlots = bookedSlots.map(slot => ({
            day: slot.day,
            time: slot.timeSlotHeader
        })).sort((a,b) => {
            const daysOrder = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
            const dayAIndex = daysOrder.indexOf(a.day);
            const dayBIndex = daysOrder.indexOf(b.day);
            if (dayAIndex !== dayBIndex) return dayAIndex - dayBIndex;
            return getTimeInMinutes(a.time) - getTimeInMinutes(b.time);
        });

        allStudents.push(studentInfo);
      });

      Logger.log("تم جلب " + allStudents.length + " طالب لصفحة 'كل الطلاب'.");
      // NEW: Ensure return value is always a valid JSON-serializable array/object
      try {
          return JSON.parse(JSON.stringify(allStudents)); // Ensure proper serialization
      } catch (e) {
          Logger.log("خطأ في تسلسل بيانات الطلاب النهائية: " + e.message);
          return { error: `خطأ داخلي في معالجة البيانات: ${e.message}` };
      }

  } catch (e) {
      Logger.log("خطأ عام في getAllStudentsData: " + e.message);
      // NEW: Ensure error object is also JSON-serializable
      try {
          return JSON.parse(JSON.stringify({ error: `خطأ أثناء جلب بيانات الطلاب: ${e.message}` }));
      } catch (jsonError) {
          return { error: `خطأ أثناء جلب بيانات الطلاب (فشل التسلسل): ${jsonError.message}` };
      }
  }
}

/**
 * دالة مساعدة لتحويل سلسلة الوقت (أو رأس الميعاد) إلى دقائق لغرض الفرز.
 * ستتعامل مع سلاسل HH:mm أو كائنات Date مباشرة.
 * @param {any} timeValue - سلسلة الوقت أو كائن Date.
 * @returns {number} الوقت بالدقائق من منتصف الليل.
 */
function getTimeInMinutes(timeValue) { // <--- هذه هي الدالة التي يجب نقلها
    // 1. التحقق أولاً من أن القيمة ليست فارغة أو غير صالحة
    if (timeValue === null || timeValue === undefined || timeValue === '') return 0;

    let hours, minutes;

    // 2. إذا كانت القيمة كائن Date (من الخلايا المنسقة كـ "وقت")
    if (timeValue instanceof Date) {
        // التحقق من صحة كائن Date وتجاهل تواريخ الصفر
        const minValidDate = new Date('1900-01-01T00:00:00.000Z');
        if (isNaN(timeValue.getTime()) || timeValue.getTime() < minValidDate.getTime()) {
            return 0; // تاريخ غير صالح أو تاريخ صفر، يعامل كـ 00:00 (0 دقيقة)
        }
        hours = timeValue.getHours();
        minutes = timeValue.getMinutes();
    } else {
        // 3. إذا لم يكن كائن Date، فافترض أنه سلسلة نصية وحاول تحليلها
        let timeString = String(timeValue).trim();

        // التعامل مع "HH:MM - HH:MM" (رؤوس الأعمدة)
        const parts = timeString.split(' - ');
        if (parts.length > 1 && parts[0].includes(':')) {
            timeString = parts[0]; // نأخذ جزء البداية فقط (HH:MM)
        }

        // التعامل مع "p X" (مسائي)
        if (timeString.toLowerCase().startsWith('p ')) {
            const hourPart = parseInt(timeString.substring(2));
            if (!isNaN(hourPart) && hourPart >= 1 && hourPart <= 12) {
                hours = (hourPart === 12) ? 12 : hourPart + 12;
                minutes = 0;
            } else return 0;
        }
        // التعامل مع "ص" (صباحي)
        else if (timeString.toLowerCase().endsWith('ص')) {
            const hourPart = parseInt(timeString.replace('ص', ''));
            if (!isNaN(hourPart) && hourPart >= 1 && hourPart <= 12) {
                hours = (hourPart === 12) ? 0 : hourPart;
                minutes = 0;
            } else return 0;
        }
        // التعامل مع "م" (مسائي)
        else if (timeString.toLowerCase().endsWith('م')) {
            const hourPart = parseInt(timeString.replace('م', ''));
            if (!isNaN(hourPart) && hourPart >= 1 && hourPart <= 12) {
                hours = (hourPart === 12) ? 12 : hourPart + 12;
                minutes = 0;
            } else return 0;
        }
        // التعامل مع التنسيق HH:MM (مثل 09:00) أو H:M (مثل 8:30)
        else if (timeString.includes(':')) {
            const timeParts = timeString.split(':').map(Number);
            if (timeParts.length === 2 && !isNaN(timeParts[0]) && !isNaN(timeParts[1])) {
                hours = timeParts[0];
                minutes = timeParts[1];
            } else return 0;
        }
        // لو الوقت رقم فقط (مثل 9 أو 10)
        else {
            const hourPart = parseInt(timeString);
            if (!isNaN(hourPart) && hourPart >= 0 && hourPart <= 23) {
                hours = hourPart;
                minutes = 0;
            } else return 0;
        }
    }

    if (isNaN(hours) || isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
        return 0;
    }
    return hours * 60 + minutes;
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
    const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون"); // شيت الطلاب التجريبيين
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

        // التحقق من عدم تسجيل الحضور مسبقًا لهذا الميعاد واليوم
        const attendanceLogData = attendanceLogSheet.getDataRange().getValues();
        for (let i = 1; i < attendanceLogData.length; i++) {
            const logRow = attendanceLogData[i];
            const logStudentID = String(logRow[1] || '').trim();
            const logTeacherID = String(logRow[2] || '').trim();
            const logDateValue = logRow[4];
            const logDate = (logDateValue instanceof Date) ? Utilities.formatDate(logDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
            const logTimeSlot = String(logRow[5] || '').trim();
            const logStatus = String(logRow[7] || '').trim();

            // إذا كان هناك سجل لنفس الطالب، ونفس المعلم، ونفس اليوم، ونفس الميعاد، فهذا تسجيل مكرر.
            if (logStudentID === studentId && logTeacherID === teacherId && logDate === todayFormatted && logTimeSlot === timeSlot) {
                return { error: `تم تسجيل الطالب ${studentId} في هذا الميعاد (${timeSlot}) وهذا اليوم (${day}) بالفعل كـ "${logStatus}". لا يمكن التسجيل مرة أخرى.` };
            }
        }

        let studentName = '';
        let isTrialStudent = studentId.startsWith("TRL"); // لتحديد ما إذا كان طالب تجريبي
        let isFreeSubscription = false; // <-- جديد: لتحديد ما إذا كان الاشتراك مجاني
        let subscriptionId = '';
        let packageName = '';
        let renewalStatus = ''; // حالة التجديد الحالية من الاشتراك
        let totalPackageSessions = 0;
        let subscriptionRowIndex = -1;
        let sessionClassType = "عادية"; // <-- جديد: لتسجيل نوع الحصة الفعلي (عادية، تجريبية، مجانية)

        if (isTrialStudent) {
            // جلب اسم الطالب التجريبي
            const trialStudentsData = trialStudentsSheet.getDataRange().getValues();
            const trialStudentRow = trialStudentsData.find(row => String(row[0] || '').trim() === studentId); // العمود A: Trial ID
            if (trialStudentRow) {
                studentName = String(trialStudentRow[1] || '').trim(); // العمود B: Student Name
                sessionClassType = "تجريبية"; // تحديد نوع الحصة كـ "تجريبية"
                Logger.log(`تسجيل حضور لطالب تجريبي: ${studentName}`);
            } else {
                throw new Error(`لم يتم العثور على اسم الطالب التجريبي بمعرف ${studentId}.`);
            }
        } else {
            // جلب اسم الطالب المشترك
            const studentsData = studentsSheet.getDataRange().getValues();
            const regularStudentRow = studentsData.find(row => String(row[0] || '').trim() === studentId); // العمود A: Student ID
            if (regularStudentRow) {
                studentName = String(regularStudentRow[1] || '').trim(); // العمود B: Student Name
                Logger.log(`تسجيل حضور لطالب مشترك: ${studentName}`);
            } else {
                throw new Error(`لم يتم العثور على اسم الطالب المشترك بمعرف ${studentId}.`);
            }

            // جلب بيانات الاشتراك (فقط للطلاب المشتركين)
            const subscriptionsData = subscriptionsSheet.getDataRange().getValues();
            for (let i = 1; i < subscriptionsData.length; i++) {
                if (String(subscriptionsData[i][1] || '').trim() === studentId) { // العمود B: Student ID
                    subscriptionRowIndex = i;
                    subscriptionId = String(subscriptionsData[i][0] || '').trim();
                    packageName = String(subscriptionsData[i][2] || '').trim();
                    renewalStatus = String(subscriptionsData[i][7] || '').trim(); // جلب الحالة الحالية
                    totalPackageSessions = getTotalSessionsForPackage(packageName);
                    
                    if (renewalStatus === "مجاني") { // <-- جديد: تحديد إذا كان الاشتراك مجاني
                        isFreeSubscription = true;
                        sessionClassType = "مجانية"; // نوع الحصة في السجل
                    }
                    break;
                }
            }
            if (subscriptionRowIndex === -1) throw new Error(`لم يتم العثور على اشتراك للطالب ID ${studentId}.`);
        }

        // تسجيل الحضور في شيت "سجل الحضور"
        const attendanceId = generateUniqueAttendanceId(attendanceLogSheet);
        attendanceLogSheet.appendRow([
            attendanceId,        // A
            studentId,           // B
            teacherId,           // C
            subscriptionId,      // D (فارغ إذا كان تجريبياً)
            today,               // E
            timeSlot,            // F
            day,                 // G
            "حضر",               // H (الحالة هنا دائماً "حضر" لهذه الدالة)
            sessionClassType,    // I (تجريبية أو عادية أو مجانية)
            ""                   // J
        ]);
        Logger.log(`تم تسجيل الحضور لـ ${studentName} (ID: ${studentId}) في ${day} ${timeSlot} كحصة ${sessionClassType}.`);

        // تحديث عدد الحصص الحاضرة في شيت "الاشتراكات الحالية" (فقط للحصص العادية والمجانية)
        if (sessionClassType === "عادية" || sessionClassType === "مجانية") {
            if (subscriptionRowIndex !== -1) { // تأكد من وجود اشتراك قبل محاولة تحديثه
                const currentAttendedSessionsCell = subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7); // العمود G
                let currentSessions = currentAttendedSessionsCell.getValue();
                currentSessions = (typeof currentSessions === 'number') ? currentSessions : 0;
                subscriptionsSheet.getRange(subscriptionRowIndex + 1, 7).setValue(currentSessions + 1);
                Logger.log(`تم تحديث عدد الحصص الحاضرة للطالب ${studentName} (ID: ${studentId}) إلى ${currentSessions + 1}.`);

                // تحديث "الحالة التفصيلية للتجديد" فقط إذا لم يكن الاشتراك مجانيًا
                if (!isFreeSubscription) { // <-- جديد: لا تغير حالة "مجاني"
                    if (totalPackageSessions > 0 && (currentSessions + 1) >= totalPackageSessions) {
                        subscriptionsSheet.getRange(subscriptionRowIndex + 1, 8).setValue("يحتاج للتجديد"); // العمود H
                        Logger.log(`حالة تجديد الطالب ${studentId} تم تحديثها إلى "يحتاج للتجديد".`);
                    }
                }
            } else {
                 Logger.log(`تحذير: لم يتم العثور على اشتراك للطالب ${studentId}. لم يتم تحديث عدد الحصص الحاضرة.`);
            }
        } else if (sessionClassType === "تجريبية") {
            // يمكن إضافة منطق لتتبع عدد الحصص التجريبية هنا إذا أردت، مثلاً في شيت "الطلاب التجريبيون"
            // (لم يتم طلب هذا التتبع حاليا، لذا نكتفي بالتسجيل في سجل الحضور)
            const trialStudentRowIndex = trialStudentsSheet.getDataRange().getValues().findIndex(row => String(row[0] || '').trim() === studentId) + 1;
            if (trialStudentRowIndex > 0) {
                // مثال: trialStudentsSheet.getRange(trialStudentRowIndex, 12).setValue((trialStudentsSheet.getRange(trialStudentRowIndex, 12).getValue() || 0) + 1);
                Logger.log(`تم تسجيل حصة تجريبية للطالب ${studentName} (ID: ${studentId}).`);
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
// 7. الدوال الخاصة بصفحة إدارة مواعيد المعلمين (Manage Slots Page)
// ==============================================================================

/**
 * تجلب جميع المواعيد من شيت "المواعيد المتاحة للمعلمين".
 * تحدد ما إذا كان الموعد متاحاً (خلية بها رأس العمود) أو محجوزاً (خلية بها ID طالب) أو غير متاح (خلية فارغة).
 *
 * @param {string} teacherName - اسم المعلم المراد جلب مواعيده.
 * @returns {Object} كائن يحتوي على { teacherName, teacherId, slots: Array<Object>, error }.
 * كل كائن slot: { dayOfWeek, colIndex, timeSlotHeader, actualSlotValue, isBooked, bookedBy: { name, id, phone } (إذا محجوز) }
 */
function getTeacherAvailableSlots(teacherName) {
  const teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const studentDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const teachersListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين"); // إضافة شيت المعلمين للحصول على ID

  if (!teacherSheet) {
    Logger.log("خطأ: لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    return { error: "شيت 'المواعيد المتاحة للمعلمين' غير موجود." };
  }
  if (!studentDataSheet) {
    Logger.log("تحذير: لم يتم العثور على شيت 'الطلاب'. قد لا تظهر تفاصيل الطلاب المحجوزة.");
  }
  if (!teachersListSheet) { // التحقق من وجود شيت المعلمين
    Logger.log("خطأ: لم يتم العثور على شيت 'المعلمين'.");
    return { error: "شيت 'المعلمين' غير موجود." };
  }

  const teacherDataRaw = teacherSheet.getDataRange().getValues(); // بيانات المواعيد للمعلمين
  if (teacherDataRaw.length < 1) {
    return { error: "شيت 'المواعيد المتاحة للمعلمين' فارغ." };
  }

  const headers = teacherDataRaw[0]; // صف العناوين
  let foundTeacherId = null;
  let teacherSlots = [];

  // --- الخطوة الأولى: جلب Teacher ID من اسم المعلم باستخدام شيت المعلمين ---
  const teachersRawData = teachersListSheet.getDataRange().getValues();
  let teacherRowInTeachersSheet = -1;
  for (let i = 1; i < teachersRawData.length; i++) {
    if (String(teachersRawData[i][1] || '').trim() === String(teacherName).trim()) { // العمود B (مؤشر 1) لاسم المعلم
      foundTeacherId = String(teachersRawData[i][0]).trim(); // العمود A (مؤشر 0) لـ Teacher ID
      teacherRowInTeachersSheet = i;
      break;
    }
  }

  if (!foundTeacherId) {
    return { error: `لم يتم العثور على المعلم "${teacherName}" في شيت المعلمين.` };
  }
  // ----------------------------------------------------------------------


  // جلب بيانات الطلاب لربط الطالب المحجوز بالـ ID بتاعه
  const studentMap = new Map();
  if (studentDataSheet) {
    const studentValues = studentDataSheet.getDataRange().getValues();
    studentValues.forEach((row, index) => {
      if (index === 0) return;
      const studentID = String(row[0] || '').trim(); // العمود A: Student ID في شيت الطلاب
      const name = String(row[1] || '').trim(); // العمود B: اسم الطالب في شيت الطلاب
      const phone = String(row[3] || '').trim(); // العمود D: رقم الهاتف (ولي الأمر) في شيت الطلاب
      if (studentID) {
        studentMap.set(studentID, { _id: studentID, name: name, phone: phone });
      }
    });
  }

  // تحديد أعمدة المواعيد ديناميكياً
  const timeSlotColumns = [];
  const startColIndexForSlots = 2; // العمود C (مؤشر 2)
  for (let i = startColIndexForSlots; i < headers.length; i++) {
    const header = String(headers[i] || '').trim();
    if (header) { // أي رأس عمود غير فارغ في هذا النطاق نعتبره ميعادًا
      timeSlotColumns.push({ index: i, header: header });
    }
  }

  if (timeSlotColumns.length === 0) {
      Logger.log("تحذير: لم يتم العثور على أي أعمدة مواعيد في شيت 'المواعيد المتاحة للمعلمين'.");
      return { error: "لم يتم تكوين جدول المواعيد في شيت المعلمين بشكل صحيح (لا توجد أعمدة مواعيد)." };
  }
  Logger.log("تم تحديد أعمدة المواعيد (الرؤوس): " + JSON.stringify(timeSlotColumns.map(c => c.header)));


  // --- الخطوة الثانية: البحث عن صفوف المعلم المحددة باستخدام Teacher ID ---
  const allTeacherRowsInSlotsSheet = teacherDataRaw.filter((row, index) => {
      if (index === 0) return false;
      // العمود A هو Teacher ID في شيت المواعيد المتاحة للمعلمين
      return String(row[0] || '').trim() === foundTeacherId;
  });

  if (allTeacherRowsInSlotsSheet.length === 0) {
      return { message: `المعلم "${teacherName}" لا يوجد لديه صفوف مواعيد مضافة في شيت 'المواعيد المتاحة للمعلمين'.` };
  }

  allTeacherRowsInSlotsSheet.forEach(row => {
      const dayOfWeek = String(row[1] || '').trim(); // العمود B: اليوم
      if (!dayOfWeek) return;

      timeSlotColumns.forEach(colInfo => {
          const slotValue = String(row[colInfo.index] || '').trim(); // قيمة الخلية
          const timeSlotHeader = colInfo.header; // رأس العمود

          const slot = {
              dayOfWeek: dayOfWeek,
              colIndex: colInfo.index,
              timeSlotHeader: timeSlotHeader,
              actualSlotValue: slotValue,
              isBooked: false,
              bookedBy: null
          };

          if (slotValue) { // لو الخلية فيها قيمة
              // إذا كانت الخلية تحتوي على Student ID (تبدأ بـ "STD" أو "p ")
              if (slotValue.startsWith("STD") || slotValue.startsWith("p ") || slotValue.startsWith("TRL")) { // أضف TRL للطلاب التجريبيين
                  const studentIdInCell = slotValue;
                  slot.isBooked = true;
                  slot.bookedBy = studentMap.get(studentIdInCell) || { _id: studentIdInCell, name: `طالب (${studentIdInCell})`, phone: "" };
                  Logger.log(`[${teacherName} - ${dayOfWeek} - ${timeSlotHeader}] محجوز بـ ID: ${studentIdInCell}`);
              }
              // إذا كانت الخلية تحتوي على رأس العمود نفسه، فهذا يعني أنها متاحة (تم تعيينها كمتاح)
              else if (slotValue === timeSlotHeader) {
                  // isBooked تظل false (الموعد متاح)
                  Logger.log(`[${teacherName} - ${dayOfWeek} - ${timeSlotHeader}] متاح (يحتوي على رأس العمود).`);
              }
              // إذا كانت الخلية تحتوي على أي نص آخر (لا يتطابق مع رأس العمود وليس ID طالب)، فهذا محجوز أيضاً (بسبب مخصص أو إجازة)
              else {
                  slot.isBooked = true;
                  slot.bookedBy = { name: slotValue, _id: "CUSTOM_BLOCKED" }; // نعتبرها محجوزة بنص الخلية
                  Logger.log(`[${teacherName} - ${dayOfWeek} - ${timeSlotHeader}] محجوز بنص: "${slotValue}"`);
              }
          }
          // إذا كانت الخلية فارغة، تظل isBooked: false (الموعد غير متاح - المعلم لا يدرس فيه)
          else {
              Logger.log(`[${teacherName} - ${dayOfWeek} - ${timeSlotHeader}] فارغ (غير متاح - معلم لا يدرس).`);
          }

          teacherSlots.push(slot);
      });
  });

  if (teacherSlots.length === 0) {
    return { teacherName: teacherName, teacherId: foundTeacherId, slots: [], message: `المعلم ${teacherName} ليس لديه مواعيد متاحة حالياً أو لم يتم تحديدها في الشيت.` };
  }

  Logger.log(`تم جلب ${teacherSlots.length} موعد للمعلم ${teacherName}.`);
  return { teacherName: teacherName, teacherId: foundTeacherId, slots: teacherSlots };
}


/**
 * تقوم بتحديث المواعيد المتاحة لمعلم معين في شيت "المواعيد المتاحة للمعلمين" ليوم معين.
 * المنطق:
 * - لو الموعد كان محجوزاً بواسطة طالب (isBooked = true) في الشيت، لن نغيره.
 * - لو الموعد تم اختياره في المودال (من خلال newSelectedSlots)، نضع رأس العمود في الخلية (متاح).
 * - لو الموعد لم يتم اختياره في المودال، وكانت الخلية لا تحتوي على حجز طالب (فارغة أو كانت رأس العمود)، نجعل الخلية فارغة (غير متاح).
 *
 * @param {string} teacherId - Teacher ID للمعلم (العمود A).
 * @param {string} day - اليوم (مثلاً: "الأحد").
 * @param {Array<Object>} newSelectedSlots - مصفوفة من كائنات المواعيد التي تم اختيارها (نريد جعلها متاحة).
 * كل كائن { dayOfWeek, timeSlotHeader }
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateTeacherSlots(teacherId, day, newSelectedSlots) {
  const teacherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!teacherSheet) {
      throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    }

    const data = teacherSheet.getDataRange().getValues();
    const headers = data[0]; // صف العناوين

    // تحديد أعمدة المواعيد ديناميكياً
    const timeSlotColMap = new Map(); // key: timeSlotHeader, value: colIndex
    const startColIndexForSlots = 2; // العمود C (مؤشر 2)
    for (let i = startColIndexForSlots; i < headers.length; i++) {
        const header = String(headers[i] || '').trim();
        if (header) {
            timeSlotColMap.set(header, i);
        }
    }

    if (timeSlotColMap.size === 0) {
        throw new Error("لم يتم العثور على أعمدة مواعيد في شيت المعلمين.");
    }

    let actualRowInSheet = -1; // رقم الصف الفعلي في الشيت (1-based)

    // البحث عن الصف المطابق للمعلم واليوم
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === String(teacherId).trim() && String(data[i][1] || '').trim() === String(day).trim()) {
            actualRowInSheet = i + 1; // +1 لتحويله إلى رقم صف حقيقي في الشيت
            break;
        }
    }

    if (actualRowInSheet === -1) {
        throw new Error(`لم يتم العثور على اليوم "${day}" للمعلم ID "${teacherId}".`);
    }
    
    // جلب الصف الحالي لليوم المحدد لتحديد المواعيد المحجوزة التي لا يجب مسحها
    const currentRowValues = teacherSheet.getRange(actualRowInSheet, 1, 1, headers.length).getValues()[0];
    
    // المواعيد التي يجب أن تظل كما هي (محجوزة بواسطة طالب أو نص مخصص)
    const protectedBookedSlots = new Set(); // key: colIndex

    timeSlotColMap.forEach((colIndex, timeSlotHeader) => {
        const currentCellValue = String(currentRowValues[colIndex] || '').trim();
        const isBookedByStudent = currentCellValue.startsWith("STD") || currentCellValue.startsWith("p ");
        const isCustomBooked = currentCellValue !== '' && currentCellValue !== timeSlotHeader && !isBookedByStudent;

        if (isBookedByStudent || isCustomBooked) {
            protectedBookedSlots.add(colIndex);
        }
    });

    const updates = []; // لتجميع التحديثات على الخلايا

    // المرور على جميع رؤوس الأعمدة (المواعيد المحتملة)
    timeSlotColMap.forEach((colIndex, timeSlotHeader) => {
        const currentCellValue = String(currentRowValues[colIndex] || '').trim();
        
        // هل هذا الموعد تم اختياره في المودال (لجعله متاحاً)؟
        const isSelectedInModal = newSelectedSlots.some(s => s.timeSlotHeader === timeSlotHeader);

        if (protectedBookedSlots.has(colIndex)) {
            // لو الموعد محجوز (بواسطة طالب أو نص مخصص)، لا تغير قيمته.
            Logger.log(`الميعاد ${timeSlotHeader} للمعلم ${teacherId} في يوم ${day} محجوز (لا تغيير).`);
        } else if (isSelectedInModal) {
            // لو الموعد لم يكن محجوزاً، وتم اختياره ليكون متاحاً (نكتب رأس العمود في الخلية)
            if (currentCellValue !== timeSlotHeader) { // فقط لو القيمة مختلفة لتجنب تحديث غير ضروري
                updates.push({ row: actualRowInSheet, col: colIndex + 1, value: timeSlotHeader });
                Logger.log(`تحديث ${timeSlotHeader} لـ ${teacherId} ${day}: أصبح متاحاً.`);
            }
        } else {
            // لو الموعد لم يكن محجوزاً، ولم يتم اختياره (نجعل الخلية فارغة - غير متاح)
            if (currentCellValue !== '') { // فقط لو الخلية مش فارغة بالفعل
                updates.push({ row: actualRowInSheet, col: colIndex + 1, value: '' });
                Logger.log(`تحديث ${timeSlotHeader} لـ ${teacherId} ${day}: أصبح غير متاح.`);
            }
        }
    });

    // تطبيق كل التحديثات مرة واحدة
    if (updates.length > 0) {
        updates.forEach(update => {
            teacherSheet.getRange(update.row, update.col).setValue(update.value);
        });
    }

    Logger.log(`تم تحديث مواعيد المعلم ${getTeacherNameById(teacherId)} ليوم ${day} بنجاح. عدد التحديثات: ${updates.length}`);
    return { success: `تم تحديث مواعيد المعلم ${getTeacherNameById(teacherId)} ليوم ${day} بنجاح.` };

  } catch (e) {
    Logger.log("خطأ في updateTeacherSlots: " + e.message);
    return { error: `فشل تحديث مواعيد المعلم: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


// ==============================================================================
// 8. الدوال الخاصة بصفحة أرشيف الطلاب (Archive Page)
// ==============================================================================

/**
 * تجلب بيانات جميع الطلاب من شيت "أرشيف الطلاب" مع تنظيف كامل للبيانات.
 *
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب المؤرشفين (صالحة للتسلسل JSON).
 * أو {Object} كائن خطأ إذا كان الشيت غير موجود.
 */
function getArchivedStudentsData() {
  const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("أرشيف الطلاب");

  if (!archiveSheet) {
    return { error: "شيت 'أرشيف الطلاب' غير موجود." };
  }

  const archivedData = archiveSheet.getDataRange().getValues();
  const archivedStudents = [];

  archivedData.forEach((row, index) => {
    if (index === 0) return; // تجاهل صف العناوين

    const isRowEmpty = row.every(cell => cell === '' || cell === null || cell === undefined);
    if (isRowEmpty) return;

    const cleanedStudent = {
      rowIndex: index + 1,
      studentID: String(row[0] ?? '').trim(),
      name: String(row[1] ?? '').trim(),
      age: isNaN(row[2]) ? '' : row[2],
      phone: String(row[3] ?? '').trim(),
      studentPhone: String(row[4] ?? '').trim(),
      country: String(row[5] ?? '').trim(),
      registrationDate: row[6] instanceof Date
        ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : '',
      basicStatus: String(row[7] ?? '').trim(),
      notes: String(row[8] ?? '').trim(),
      archivedDate: row[9] instanceof Date
        ? Utilities.formatDate(row[9], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : '',
      archiveReason: String(row[10] ?? '').trim(),
      archiveSubscriptionDetails: String(row[11] ?? '').trim()
    };

    archivedStudents.push(cleanedStudent);
  });

  Logger.log("✅ تم جلب " + archivedStudents.length + " طالب من الأرشيف.");
  Logger.log("🔎 عينة من أول طالب:", JSON.stringify(archivedStudents[0]));

  return archivedStudents;
}






/**
 * تعيد تفعيل طالب من الأرشيف إلى شيت "الطلاب".
 *
 * @param {string} studentID - معرف الطالب المراد إعادة تفعيله.
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function reactivateStudentFromArchive(studentID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName("الطلاب");
  const archiveSheet = ss.getSheetByName("أرشيف الطلاب");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!archiveSheet) throw new Error("شيت 'أرشيف الطلاب' غير موجود.");

    const archivedData = archiveSheet.getDataRange().getValues();
    let archiveRowIndex = -1;
    let row = null;

    for (let i = 1; i < archivedData.length; i++) {
      if (String(archivedData[i][0]).trim() === String(studentID).trim()) {
        archiveRowIndex = i;
        row = archivedData[i];
        break;
      }
    }

    if (!row) {
      throw new Error(`لم يتم العثور على الطالب ID ${studentID} في الأرشيف.`);
    }

    // إعداد بيانات إعادة التفعيل بدقة
    const studentRow = [
      row[0] || '', // Student ID
      row[1] || '', // الاسم
       '', // السن
      row[2] || '', // رقم الهاتف (ولي الأمر)
      '', // رقم هاتف الطالب
       '', // البلد
      row[5] || '', // تاريخ التسجيل
      "معلق",       // الحالة
      row[8] || ''  // الملاحظات
    ];

    // إضافة البيانات للسطر التالي المتاح
    studentsSheet.appendRow(studentRow);
    Logger.log(`✅ تمت إعادة الطالب ${row[1]} إلى شيت 'الطلاب'.`);

    // حذف الطالب من الأرشيف
    archiveSheet.deleteRow(archiveRowIndex + 1);
    Logger.log(`🗑️ تم حذف الطالب ${row[1]} من الأرشيف.`);

    return { success: `تمت إعادة تفعيل الطالب ${row[1]} بنجاح.` };

  } catch (e) {
    Logger.log("خطأ في reactivateStudentFromArchive: " + e.message);
    return { error: `فشل إعادة التفعيل: ${e.message}` };
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
        if (packageName === "اشتراك نصف سنوي" || packageName === "مخصص" || currentRenewalStatus === "تجريبي"|| currentRenewalStatus === "مجاني") {
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
        }else if (currentRenewalStatus === "مجاني") { // <-- جديد: إذا كان مجاني، يبقى مجاني
             newRenewalStatus = "مجاني";
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
  let studentBasicInfo = null; // تعريف متغير studentBasicInfo

  // 1. البحث عن الطالب في شيت "الطلاب" أولاً
  for (let i = 1; i < studentData.length; i++) {
    if (String(studentData[i][0] || '').trim() === String(studentId).trim()) { // العمود A: Student ID
      studentBasicInfo = {
        rowIndex: i + 1, // الصف في الشيت (1-based)
        studentID: String(studentData[i][0] || '').trim(),
        name: String(studentData[i][1] || '').trim(),
        age: studentData[i][2],
        phone: String(studentData[i][3] || '').trim(), // رقم الهاتف ولي الأمر
        studentPhone: String(studentData[i][4] || '').trim(), // رقم هاتف الطالب (إن وجد)
        country: Object.prototype.toString.call(studentData[i][5]) === '[object Date]'
          ? Utilities.formatDate(studentData[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd")
          : String(studentData[i][5] || '').trim(),
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
        studentFound.subscriptionType = String(subscriptionsValues[i][12] || '').trim(); // نوع الاشتراك (M)
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
  const studentBookedSlotsMapForSingleStudent = new Map(); // key: Student ID, value: [{day, timeSlotHeader, teacherId}, ...]
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
                if (!studentBookedSlotsMapForSingleStudent.has(studentId)) {
                    studentBookedSlotsMapForSingleStudent.set(studentId, []);
                }
                studentBookedSlotsMapForSingleStudent.get(studentId).push({
                    teacherId: teacherIdInSlot,
                    day: dayOfWeek,
                    timeSlotHeader: timeSlotHeader,
                    actualSlotValue: slotValue // القيمة الفعلية في الخلية (Student ID)
                });
            }
        });
    }
  }

  // هذه هي النقطة الرئيسية للتعديل: سنضيف مصفوفة بكل المواعيد المحجوزة
  const bookedSlots = studentBookedSlotsMapForSingleStudent.get(studentId) || [];
  studentFound.allBookedScheduleSlots = bookedSlots.map(slot => ({
      day: slot.day,
      time: slot.timeSlotHeader // الميعاد برأس العمود
  })).sort((a,b) => { // فرز المواعيد لضمان عرضها بشكل مرتب
      const daysOrder = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
      const dayAIndex = daysOrder.indexOf(a.day);
      const dayBIndex = daysOrder.indexOf(b.day);
      if (dayAIndex !== dayBIndex) return dayAIndex - dayBIndex;
      return getTimeInMinutes(a.time) - getTimeInMinutes(b.time);
  });

  // لا نحتاج لـ day1, time1, day2, time2 بشكل مباشر في الواجهة الأمامية بعد الآن
  // لكن يمكن إبقاؤها هنا للتوافق مع الدوال التي قد تعتمد عليها بشكل ضمني
  // ومع ذلك، دالة fillStudentForm في الواجهة الأمامية ستعتمد على allBookedScheduleSlots مباشرة.
  studentFound.day1 = studentFound.allBookedScheduleSlots[0] ? studentFound.allBookedScheduleSlots[0].day : '';
  studentFound.time1 = studentFound.allBookedScheduleSlots[0] ? studentFound.allBookedScheduleSlots[0].time : '';
  studentFound.day2 = studentFound.allBookedScheduleSlots[1] ? studentFound.allBookedScheduleSlots[1].day : '';
  studentFound.time2 = studentFound.allBookedScheduleSlots[1] ? studentFound.allBookedScheduleSlots[1].time : '';


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
 * والمتوقع أن يحتوي على: { studentID, rowIndex, editName, editAge, editPhone, editTeacher,
 * editedSlots: Array<{day: string, time: string}>, editPaymentStatus, editSubscriptionPackage, editSubscriptionType }
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function updateStudentDataWithReassignment(updatedData) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet) throw new Error("شيت 'الطلاب' غير موجود.");
    if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    if (!teachersSheet) throw new Error("لم يتم العثور على شيت 'المعلمين'.");
    if (!packagesSheet) throw new Error("لم يتم العثور على شيت 'الباقات'.");

    const studentId = updatedData.studentID;
    const studentRowInStudentsSheet = updatedData.rowIndex; // صف الطالب في شيت "الطلاب"
    if (!studentId || !studentRowInStudentsSheet) {
      throw new Error("بيانات الطالب (Student ID أو رقم الصف) غير كاملة للتحديث.");
    }

    // 1. تحديث بيانات الطالب في شيت "الطلاب"
    // الأعمدة: Student ID(A), اسم الطالب(B), السن(C), رقم الهاتف(D), رقم هاتف الطالب(E), البلد(F), تاريخ التسجيل(G), الحالة الأساسية للطالب(H), ملاحظات(I)
    studentsSheet.getRange(studentRowInStudentsSheet, 2).setValue(updatedData.editName); // اسم الطالب (B)
    studentsSheet.getRange(studentRowInStudentsSheet, 3).setValue(updatedData.editAge);    // السن (C)
    studentsSheet.getRange(studentRowInStudentsSheet, 4).setValue(updatedData.editPhone); // رقم الهاتف (D)

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

    const newPackageName = updatedData.editSubscriptionPackage;
    const newPackageDetails = getPackageDetails(newPackageName);
    let newSessionDuration = 0;
    if (newPackageDetails) {
        newSessionDuration = newPackageDetails['مدة الحصة (دقيقة)'] || 0;
    }


    subscriptionsSheet.getRange(subscriptionRowInSheet, 3).setValue(newPackageName); // اسم الباقة (C)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 4).setValue(newTeacherId);                       // Teacher ID (D)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 8).setValue(updatedData.editPaymentStatus);      // الحالة التفصيلية للتجديد (H)
    subscriptionsSheet.getRange(subscriptionRowInSheet, 13).setValue(updatedData.editSubscriptionType);  // نوع الاشتراك (M)

    // 3. معالجة إعادة تخصيص المواعيد في شيت "المواعيد المتاحة للمعلمين"
    // أ. تحرير جميع المواعيد التي حجزها هذا الطالب حالياً
    // بما أننا سنعيد حجز المواعيد ديناميكياً بناءً على editedSlots، يجب أن نحرر كل المواعيد المحجوزة
    // للطالب أولاً، بما في ذلك المواعيد المزدوجة لحصص الساعة.
    const oldBookedSlotsByStudent = getBookedSlotsByStudentId(teachersAvailableSlotsSheet, studentId);
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
        }
    });

    // ب. حجز المواعيد الجديدة
    const newBookedSlotsToAssign = updatedData.editedSlots || [];
    let bookingType = (updatedData.editPaymentStatus === "تجريبي") ? "تجريبي" : "عادي";

    for (let i = 0; i < newBookedSlotsToAssign.length; i++) {
        const slot = newBookedSlotsToAssign[i];
        if (slot.day && slot.time) {
            const initialTimeSlotHeader = slot.time;
            if (initialTimeSlotHeader) {
                // **أولاً: حجز الموعد الأساسي الذي تم اختياره من قبل المستخدم**
                let bookResult = bookTeacherSlot(
                    teachersAvailableSlotsSheet,
                    newTeacherId,
                    slot.day,
                    initialTimeSlotHeader,
                    studentId,
                    bookingType
                );
                if (bookResult.error) {
                    Logger.log(`خطأ في حجز الميعاد الجديد ${slot.day} ${initialTimeSlotHeader}: ${bookResult.error}`);
                    throw new Error(`تعذر حجز الميعاد الجديد ${slot.day} ${initialTimeSlotHeader}: ${bookResult.error}`);
                } else {
                    Logger.log(`تم حجز الميعاد الجديد ${slot.day} ${initialTimeSlotHeader} للطالب ${studentId}.`);
                }

                // **ثانياً: إذا كانت مدة الحصة ساعة (60 دقيقة) للباقة الجديدة، احجز الموعد التالي**
                if (newSessionDuration === 60) {
                    const nextTimeSlotHeader = getNextTimeSlotHeader(initialTimeSlotHeader); // نحسب الميعاد التالي من الميعاد الأساسي المحجوز
                    if (nextTimeSlotHeader) {
                        bookResult = bookTeacherSlot( // إعادة استخدام bookResult
                            teachersAvailableSlotsSheet,
                            newTeacherId,
                            slot.day,
                            nextTimeSlotHeader,
                            studentId,
                            bookingType
                        );
                        if (bookResult.error) {
                            Logger.log(`خطأ في حجز الميعاد التالي ${slot.day} ${nextTimeSlotHeader} للطالب ${studentId}: ${bookResult.error}`);
                            throw new Error(`تعذر حجز الميعاد التالي ${slot.day} ${nextTimeSlotHeader}: ${bookResult.error}`);
                        } else {
                            Logger.log(`تم حجز الميعاد التالي ${slot.day} ${nextTimeSlotHeader} للطالب ${studentId}.`);
                        }
                    } else {
                        Logger.log(`تحذير: لم يتم العثور على ميعاد تالي صالح للميعاد ${slot.day} ${initialTimeSlotHeader} لحصة مدتها ساعة للباقة الجديدة.`);
                    }
                }
            } else {
                Logger.log(`تحذير: تنسيق وقت غير صالح للميعاد الجديد: '${slot.time}'. تم تخطي حجز هذا الميعاد.`);
            }
        }
    }

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
                subscriptionType: String(data[i][12] || '').trim()
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
        targetCell.setValue(timeSlotHeader);
        Logger.log(`تم تحرير الميعاد ${day} ${timeSlotHeader} للمعلم ${teacherId} من الطالب ${expectedStudentId}.`);
        return { success: 'تم تحرير الميعاد بنجاح.' };
    } else {
        Logger.log(`الميعاد ${day} ${timeSlotHeader} للمعلم ${teacherId} ليس محجوزاً بواسطة الطالب ${expectedStudentId} (قيمته: ${currentCellValue}).`);
        return { error: 'الميعاد ليس محجوزاً بهذا الطالب أو محجوز بنص آخر.' };
    }
}


/**
 * حذف طالب ونقله إلى الأرشيف بتنسيق الأعمدة الصحيح.
 * @param {Object} studentInfo - يحتوي على studentID و rowIndex و name.
 * @returns {Object} نتيجة العملية.
 */
function deleteStudentAndArchive(studentInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = ss.getSheetByName("الطلاب");
  const subscriptionsSheet = ss.getSheetByName("الاشتراكات الحالية");
  const teachersAvailableSlotsSheet = ss.getSheetByName("المواعيد المتاحة للمعلمين");
  const archiveSheet = ss.getSheetByName("أرشيف الطلاب");
  const teachersSheet = ss.getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!studentsSheet || !subscriptionsSheet || !teachersAvailableSlotsSheet || !teachersSheet || !archiveSheet) {
      throw new Error("شيت مطلوب غير موجود.");
    }

    const studentID = studentInfo.studentID;
    const rowIndex = studentInfo.rowIndex;
    const studentRow = studentsSheet.getRange(rowIndex, 1, 1, 9).getValues()[0];

    const studentName = studentRow[1] || '';
    const guardianPhone = studentRow[3] || '';
    const notes = studentRow[8] || '';

    // تحرير المواعيد
    const bookedSlots = getBookedSlotsByStudentId(teachersAvailableSlotsSheet, studentID);
    bookedSlots.forEach(slot => {
      releaseTeacherSlot(teachersAvailableSlotsSheet, slot.teacherId, slot.day, slot.timeSlotHeader, studentID);
    });

    // الاشتراك
    let teacherId = '';
    let teacherName = '';
    let packageName = '';
    let subscriptionDetails = 'لا توجد تفاصيل اشتراك.';

    const subsData = subscriptionsSheet.getDataRange().getValues();
    for (let i = 1; i < subsData.length; i++) {
      if (subsData[i][1] === studentID) {
        packageName = subsData[i][2] || '';
        teacherId = subsData[i][3] || '';
        subscriptionsSheet.deleteRow(i + 1);
        break;
      }
    }

    if (teacherId) {
      const teachersData = teachersSheet.getDataRange().getValues();
      for (let j = 1; j < teachersData.length; j++) {
        if (teachersData[j][0] === teacherId) {
          teacherName = teachersData[j][1] || '';
          break;
        }
      }
    }

    subscriptionDetails = `Teacher ID: ${teacherId}, Teacher Name: ${teacherName}, Package: ${packageName}`;

    const archiveDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const reason = `تم حذف الطالب (${studentName}) بواسطة المشرف`;

    // ترتيب البيانات حسب الأعمدة المطلوبة
    archiveSheet.appendRow([
      studentID,                  // A - Student ID
      studentName,               // B - اسم الطالب
      guardianPhone,             // C - رقم الهاتف (ولي الأمر)
      teacherName,               // D - اسم المعلم وقت الأرشفة
      packageName,               // E - اسم الباقة وقت الأرشفة
      archiveDate,               // F - تاريخ الأرشفة
      reason,                    // G - سبب الأرشفة
      notes                      // H - ملاحظات الأرشفة
    ]);

    studentsSheet.deleteRow(rowIndex);

    return { success: `✅ تم حذف الطالب ${studentName} ونقله إلى الأرشيف.` };

  } catch (e) {
    return { error: `❌ فشل الحذف: ${e.message}` };
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
  const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!trialStudentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب التجريبيون'.");
    if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
    if (!teachersSheet) throw new Error("لم يتم العثور على شيت 'المعلمين'.");

    // 1. جلب Teacher ID من اسم المعلم
    const teacherId = getTeacherIdByName(formData.trialRegTeacher); // دالة مساعدة موجودة
    if (!teacherId) throw new Error(`لم يتم العثور على Teacher ID للمعلم: ${formData.trialRegTeacher}`);
    const teacherName = getTeacherNameById(teacherId); // دالة مساعدة موجودة

    // 2. حفظ الطالب في شيت "الطلاب التجريبيون"
    const newTrialId = generateUniqueTrialId(trialStudentsSheet); // دالة جديدة
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // الأعمدة في شيت "الطلاب التجريبيون": Trial ID, Student Name, Age, Phone Number, Teacher ID, Teacher Name, Day, Time, Registration Date, Notes, Status
    trialStudentsSheet.appendRow([
      newTrialId,             // Trial ID (A)
      formData.trialRegName,  // Student Name (B)
      formData.trialRegAge,   // Age (C)
      String(formData.trialRegPhone).trim(), // Phone Number (D)
      teacherId,              // Teacher ID (E)
      teacherName,            // Teacher Name (F)
      formData.trialRegDay1,  // Day (G)
      formData.trialRegTime1, // Time (H)
      today,                  // Registration Date (I)
      "",                     // Notes (J)
      "تجريبي"                // Status (K)
    ]);
    Logger.log(`تم حفظ الطالب التجريبي ${formData.trialRegName} (ID: ${newTrialId}) في شيت 'الطلاب التجريبيون'.`);

    // 3. حجز الميعاد في شيت "المواعيد المتاحة للمعلمين"
    // نستخدم ID الطالب التجريبي لتمييز الحجز
    const bookingTypeForSlot = "تجريبي";
    const result = bookTeacherSlot(
        teachersAvailableSlotsSheet,
        teacherId,
        formData.trialRegDay1,
        formData.trialRegTime1,
        newTrialId, // استخدم Trial ID هنا
        bookingTypeForSlot
    );
    if (result.error) {
        Logger.log(`خطأ في حجز الميعاد ${formData.trialRegDay1} ${formData.trialRegTime1} للطالب التجريبي ${newTrialId}: ${result.error}`);
        throw new Error(`تعذر حجز الميعاد التجريبي ${formData.trialRegDay1} ${formData.trialRegTime1}: ${result.error}`);
    } else {
        Logger.log(`تم حجز الميعاد ${formData.trialRegDay1} ${formData.trialRegTime1} للطالب التجريبي ${newTrialId}.`);
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
 * @param {string} selectedPackageName - اسم الباقة التي اختارها الطالب.
 * @param {string} paymentStatus - حالة الدفع الأولية للاشتراك الجديد.
 * @param {string} subscriptionType - نوع الاشتراك الذي اختاره الطالب.
 * @param {string} oldTrialDay - يوم الميعاد التجريبي الأصلي. // NEW PARAMETER
 * @param {string} oldTrialTime - ميعاد الميعاد التجريبي الأصلي. // NEW PARAMETER
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ.
 */
function convertTrialToRegistered(trialStudentId, selectedPackageName, paymentStatus, subscriptionType, oldTrialDay, oldTrialTime) { // NEW PARAMETERS
    const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
    const trialStudentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب التجريبيون");
    const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
    const teachersAvailableSlotsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المواعيد المتاحة للمعلمين");
    const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);

        if (!studentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب'.");
        if (!trialStudentsSheet) throw new Error("لم يتم العثور على شيت 'الطلاب التجريبيون'.");
        if (!subscriptionsSheet) throw new Error("لم يتم العثور على شيت 'الاشتراكات الحالية'.");
        if (!teachersAvailableSlotsSheet) throw new Error("لم يتم العثور على شيت 'المواعيد المتاحة للمعلمين'.");
        if (!packagesSheet) throw new Error("لم يتم العثور على شيت 'الباقات'.");

        // 1. جلب بيانات الطالب التجريبي (فقط للتأكد من وجوده وجلب المعرفات الأخرى)
        const trialStudentsData = trialStudentsSheet.getDataRange().getValues();
        let trialStudentRowIndex = -1;
        let trialStudentData = null; // لم نعد نستخدم day1 و time1 من هذا الكائن مباشرة للحجز

        for (let i = 1; i < trialStudentsData.length; i++) {
            if (String(trialStudentsData[i][0] || '').trim() === String(trialStudentId).trim()) {
                trialStudentRowIndex = i;
                trialStudentData = trialStudentsData[i]; // فقط لجلب teacherId
                break;
            }
        }

        if (trialStudentRowIndex === -1 || !trialStudentData) {
            throw new Error(`لم يتم العثور على الطالب التجريبي بمعرف ${trialStudentId}.`);
        }

        const name = String(trialStudentData[1] || '').trim();
        const teacherId = String(trialStudentData[4] || '').trim(); // المعلم الذي كان يدرس له الحصة التجريبية


        // 2. حفظ الطالب في شيت "الطلاب" (كطالب مشترك الآن)
        const newStudentId = generateUniqueStudentId(studentsSheet);
        const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

        let studentBasicStatus = "مشترك";

        studentsSheet.appendRow([
            newStudentId,
            name,
            trialStudentData[2], // Age
            trialStudentData[3], // Phone
            "", "", today,
            studentBasicStatus,
            "تم التحويل من تجريبي"
        ]);
        Logger.log(`تم حفظ الطالب ${name} (ID: ${newStudentId}) في شيت 'الطلاب'.`);

        // 3. إنشاء اشتراك في شيت "الاشتراكات الحالية"
        const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet);
        const packageDetails = getPackageDetails(selectedPackageName);

        let subscriptionRenewalStatus = "تم التجديد";
        let totalClassesAttended = 0;
        let sessionDuration = 0; // مدة الحصة بالدقائق للباقة الجديدة

        if (packageDetails) {
            sessionDuration = packageDetails['مدة الحصة (دقيقة)'] || 0;
            // ... بقية المنطق الخاص بتفاصيل الباقة والمبالغ
        }

        // ... (بقية كود إنشاء الاشتراك في الاشتراكات الحالية) ...

        subscriptionsSheet.appendRow([
            newSubscriptionId,
            newStudentId,
            selectedPackageName,
            teacherId,
            today,
            // ... rest of subscription fields
            totalClassesAttended,
            subscriptionRenewalStatus,
            today,
            (packageDetails ? packageDetails['السعر'] : 0),
            (paymentStatus === "تم الدفع" ? (packageDetails ? packageDetails['السعر'] : 0) : 0),
            (paymentStatus === "تم الدفع" ? 0 : (packageDetails ? packageDetails['السعر'] : 0)),
            subscriptionType,
            "تم التحويل من طالب تجريبي"
        ]);
        Logger.log(`تم إنشاء اشتراك (${newSubscriptionId}) للطالب ${newStudentId} في شيت 'الاشتراكات الحالية'.`);

        // 4. تحديث المواعيد في شيت "المواعيد المتاحة للمعلمين"
        // أ. تحرير الميعاد القديم المحجوز بمعرف الطالب التجريبي
        // نستخدم oldTrialDay و oldTrialTime مباشرة من البارامترات
        const trialOldTimeHeaderToRelease = oldTrialTime;

        const releaseResult = releaseTeacherSlot(
            teachersAvailableSlotsSheet,
            teacherId, // معلم الحصة التجريبية
            oldTrialDay, // يوم الحصة التجريبية الأصلية
            trialOldTimeHeaderToRelease, // ميعاد الحصة التجريبية الأصلية
            trialStudentId // معرف الطالب التجريبي
        );
        if (releaseResult.error) {
            Logger.log(`تحذير: خطأ في تحرير الميعاد القديم للطالب التجريبي ${trialStudentId}: ${releaseResult.error}`);
        } else {
            Logger.log(`تم تحرير الميعاد القديم للطالب التجريبي ${trialStudentId}.`);
        }

        // ب. حجز المواعيد الجديدة للطالب المشترك
        const newStudentBookingType = (paymentStatus === "حلقة تجريبية") ? "تجريبي" : "عادي"; // نوع الحجز

        // الميعاد الأول هو الميعاد التجريبي الأصلي
        const initialSlotHeader = oldTrialTime;

        let bookResult = bookTeacherSlot(
            teachersAvailableSlotsSheet,
            teacherId,
            oldTrialDay,
            initialSlotHeader,
            newStudentId,
            newStudentBookingType
        );
        if (bookResult.error) {
            Logger.log(`خطأ في حجز الميعاد الأول للطالب المشترك ${newStudentId}: ${bookResult.error}`);
            throw new Error(`تعذر حجز الميعاد الأول للطالب المشترك: ${bookResult.error}`);
        } else {
            Logger.log(`تم حجز الميعاد الأول للطالب المشترك ${newStudentId}.`);
        }

        // إذا كانت الباقة الجديدة مدتها ساعة، احجز الميعاد التالي
        if (sessionDuration === 60) {
            const nextTimeSlotHeader = getNextTimeSlotHeader(initialSlotHeader);
            if (nextTimeSlotHeader) {
                bookResult = bookTeacherSlot(
                    teachersAvailableSlotsSheet,
                    teacherId,
                    oldTrialDay,
                    nextTimeSlotHeader,
                    newStudentId,
                    newStudentBookingType
                );
                if (bookResult.error) {
                    Logger.log(`خطأ في حجز الميعاد التالي للطالب المشترك ${newStudentId}: ${bookResult.error}`);
                    throw new Error(`تعذر حجز الميعاد التالي للطالب المشترك: ${bookResult.error}`);
                } else {
                    Logger.log(`تم حجز الميعاد التالي للطالب المشترك ${newStudentId}.`);
                }
            } else {
                Logger.log(`تحذير: لم يتم العثور على ميعاد تالي صالح للميعاد ${oldTrialDay} ${initialSlotHeader} لحصة مدتها ساعة.`);
            }
        }


        // 5. حذف الطالب من شيت "الطلاب التجريبيون"
        trialStudentsSheet.deleteRow(trialStudentRowIndex + 1);
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
     let isFreeSubscription = false;

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
 * دالة مساعدة للتحقق مما إذا كانت السلسلة النصية تحتوي على أحرف عربية.
 * @param {string} text - السلسلة النصية المراد فحصها.
 * @returns {boolean} - true إذا كانت تحتوي على أحرف عربية، false بخلاف ذلك.
 */
function containsArabic(text) {
  if (typeof text !== 'string') return false;
  // هذا التعبير النمطي يطابق نطاقات Unicode الخاصة بالأحرف العربية
  const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDCF\uFDF0-\uFDFF\uFE70-\uFEFF]/;
  return arabicRegex.test(text);
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
      const rawPaymentDateValue = paymentRecordsData[i][3]; // الحصول على القيمة الخام من الشيت

      let formattedPaymentDate;
      // التحقق مما إذا كانت القيمة هي كائن Date صالح
      if (rawPaymentDateValue instanceof Date && !isNaN(rawPaymentDateValue.getTime())) {
        // إذا كانت كذلك، قم بتنسيقها كتاريخ (yyyy-MM-dd)
        formattedPaymentDate = Utilities.formatDate(rawPaymentDateValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        // إذا لم تكن كائن Date صالح (أي أنها نص، أو Date object غير صالح)، قم بتحويلها إلى نص نظيف
        formattedPaymentDate = String(rawPaymentDateValue || '').trim();
      }

      if (String(paymentRecordsData[i][2] || '').trim() === String(studentID).trim()) { // العمود C: Student ID
        studentPayments.push({
          paymentId: paymentRecordsData[i][0],
          subscriptionId: paymentRecordsData[i][1],
          studentId: paymentRecordsData[i][2],
          paymentDate: formattedPaymentDate, // استخدام التاريخ المنسق أو النصي
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
    }else if (renewalData.paymentStatus === "مجاني") { // <-- جديد
        paidAmount = 0;
        remainingAmount = 0;
        newRenewalStatus = "مجاني"; // حالة تجديد خاصة للمجاني
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
    if (renewalData.paymentStatus === "تم الدفع" || renewalData.paymentStatus === "مجاني") {
        const studentsData = studentsSheet.getDataRange().getValues();
        const studentRow = studentsData.find(row => String(row[0] || '').trim() === renewalData.studentID);
        if (studentRow) {
            const studentSheetRowIndex = studentsData.indexOf(studentRow) + 1;
            studentsSheet.getRange(studentSheetRowIndex, 8).setValue("مشترك"); // العمود H: الحالة الأساسية للطالب
            Logger.log(`تم تحديث الحالة الأساسية للطالب ${renewalData.studentID} إلى "مشترك".`);
        }
    }


    // تسجيل الدفعة المرتبطة بالتجديد (إذا تم الدفع)
    if (renewalData.paymentStatus !== "لم يتم الدفع" && renewalData.paymentStatus !== "مجاني") {
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
 * @param {any} timeValue - قيمة الوقت (يمكن أن تكون سلسلة نصية أو كائن Date أو فارغة).
 * @returns {string} الميعاد بتنسيق رأس العمود (مثلاً "09:00 - 09:30").
 */
function convertOldPlainTimeFormatToHeaderFormat(timeValue) {
  // 1. التحقق أولاً من أن القيمة ليست فارغة أو غير صالحة
  if (timeValue === null || timeValue === undefined || timeValue === '') {
    return ''; // ارجع فارغًا مباشرة إذا كانت القيمة فارغة
  }

  let hours, minutes;

  // 2. معالجة كائن Date (من الخلايا المنسقة كـ "وقت" في الشيت القديم)
  if (timeValue instanceof Date) {
    if (isNaN(timeValue.getTime())) return '';

    const isZeroDate = timeValue.getFullYear() === 1899 && timeValue.getMonth() === 11 && timeValue.getDate() === 30;

    // إذا لم يكن وقتًا فقط وكان التاريخ قديم جدًا → تجاهله
    if (!isZeroDate && timeValue.getTime() < new Date('1900-01-01T00:00:00.000Z').getTime()) {
      return '';
    }

    hours = timeValue.getHours();
    minutes = timeValue.getMinutes();
  }
  // 3. معالجة السلاسل النصية
  else {
    let timeString = String(timeValue).trim();

    // تطبيع الأرقام العربية إلى هندية (اختياري لكنه مفيد)
    timeString = timeString.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));

    // توحيد التعبيرات وإزالة الزوائد
    timeString = timeString.replace(':00', '');
    timeString = timeString.replace(/\s+/g, '');
    timeString = timeString.replace('ظ', 'ص');

    const ampmPattern = /(\d{1,2}(:\d{2})?)([صم])/; // 9ص, 9:30ص, 9م, 9:30م
    const basic24hrPattern = /(\d{1,2}:\d{2})/;     // 18:00, 9:00
    const hourOnlyPattern = /^(\d{1,2})$/;           // 9, 10

    let match;

    if (match = timeString.match(ampmPattern)) {
      let hourPart = parseInt(match[1].split(':')[0]);
      let minutePart = match[2] ? parseInt(match[2].substring(1)) : 0;
      const ampm = match[3];

      if (ampm === 'م' && hourPart !== 12) {
        hours = hourPart + 12;
      } else if (ampm === 'ص' && hourPart === 12) {
        hours = 0;
      } else {
        hours = hourPart;
      }
      minutes = minutePart;
    }
    else if (match = timeString.match(basic24hrPattern)) {
      const timeParts = match[1].split(':').map(Number);
      hours = timeParts[0];
      minutes = timeParts[1];
    }
    else if (match = timeString.match(hourOnlyPattern)) {
      const hourPart = parseInt(match[1]);
      if (!isNaN(hourPart) && hourPart >= 0 && hourPart <= 23) {
        hours = hourPart;
        minutes = 0;
      } else return '';
    }
    else {
      return '';
    }
  }

  // التحقق من صحة الساعات والدقائق النهائية
  if (isNaN(hours) || isNaN(minutes) || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
    return '';
  }

  // بناء الجزء الأول من رأس العمود (HH:mm)
  const startTime24hr = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;

  // حساب وقت النهاية (+30 دقيقة)
  const startDate = new Date();
  startDate.setHours(hours, minutes, 0, 0);

  const endDate = new Date(startDate.getTime() + 30 * 60 * 1000); // إضافة 30 دقيقة

  const endHours = endDate.getHours();
  const endMinutes = endDate.getMinutes();

  const endTime24hr = `${String(endHours).padStart(2, '0')}:${String(endMinutes).padStart(2, '0')}`;

  // بناء سلسلة رأس العمود
  return `${startTime24hr} - ${endTime24hr}`;
}





/**
 * دالة مساعدة لاستخلاص الوقت بتنسيق HH:mm من كائن Date، مع معالجة التواريخ غير الصالحة أو "تاريخ الصفر".
 *
 * @param {Date} dateObject - كائن Date.
 * @returns {string} الوقت بتنسيق HH:mm (مثلاً "18:30")، أو "00:00" إذا كان كائن Date غير صالح.
 */
function formatDateObjectToHHMM(dateObject) {
  // التحقق أولاً من أن القيمة كائن Date وصالحة
  if (!(dateObject instanceof Date) || isNaN(dateObject.getTime())) {
    return '00:00'; // ليس كائن Date أو تاريخ غير صالح
  }

  // التحقق مما إذا كان كائن Date يمثل تاريخ صفر (مثل 1899-12-30)
  // وهو التاريخ الذي يُستخدم لتمثيل قيم الوقت فقط أو الخلايا الفارغة المنسقة كـ Date
  const minValidDate = new Date('1900-01-01T00:00:00.000Z'); // استخدام ISO format لضمان التوقيت العالمي

  if (dateObject.getTime() < minValidDate.getTime()) {
    Logger.log("Info: formatDateObjectToHHMM received a 'zero' or very old Date object, treating as 00:00: " + dateObject);
    return '00:00'; // تاريخ قديم جدًا، عاملها كـ "00:00"
  }

  // استخدام Utilities.formatDate لتنسيق الوقت من كائن Date مباشرةً
  return Utilities.formatDate(dateObject, Session.getScriptTimeZone(), "HH:mm");
}


// ==============================================================================
// 10. الدوال الخاصة بنقل البيانات (Migration Functions)
// ==============================================================================

/**
 * دالة رئيسية لنقل بيانات الطلاب من شيتات الملف القديم
 * إلى هيكل الشيتات الجديد.
 *
 * ستجلب البيانات من:
 * 1. الملف القديم، شيت "بيانات الطلبة": معرف الطالب (A), اسم الطالب (B), السن (C), رقم الهاتف (D), نوع الاشتراك (G), اسم الباقة (H), عدد الحصص الحاضرة (I).
 * 2. الملف القديم، شيت "مواعيد الطلبة": معرف الطالب (A), اسم المعلم (E), تاريخ بداية الاشتراك (L).
 * 3. الملف القديم، شيت "المعلمين": اسم المعلم (B), معرف المعلم (A).
 *
 * ثم تضيفها إلى:
 * 1. شيت "الطلاب" الجديد.
 * 2. شيت "الاشتراكات الحالية" الجديد.
 *
 * لن تقوم بحجز المواعيد في شيت "المواعيد المتاحة للمعلمين" الجديد تلقائياً.
 *
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ وملخص.
 */
function migrateStudentsOnly() {
  // === قم بتغيير هذا الـ ID بالـ ID الخاص بملف Google Sheets القديم ===
  const OLD_SPREADSHEET_ID = "1XWRFKp-LM7detp42X4bqjhVgkaJT00of6FWvtx8sYL0"; // <--- تأكد من تحديث هذا الـ ID
  // ====================================================================

  let oldStudentsSpreadsheet;
  try {
    oldStudentsSpreadsheet = SpreadsheetApp.openById(OLD_SPREADSHEET_ID);
  } catch (e) {
    return { error: `فشل فتح ملف Google Sheets القديم بالـ ID المقدم: ${OLD_SPREADSHEET_ID}. الخطأ: ${e.message}` };
  }

  const oldStudentInfoSheet = oldStudentsSpreadsheet.getSheetByName("بيانات الطلبة"); // العمود A: معرف الطالب, B: اسم الطالب, C: السن, D: رقم الهاتف, G: نوع الاشتراك, H: اسم الباقة, I: عدد الحصص الحاضرة
  const oldStudentScheduleSheet = oldStudentsSpreadsheet.getSheetByName("مواعيد الطلبة"); // العمود A: معرف الطالب, E: اسم المعلم, L: تاريخ بداية الاشتراك
  const oldTeachersSheet = oldStudentsSpreadsheet.getSheetByName("المعلمين"); // العمود B: اسم المعلم, A: معرف المعلم

  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب");
  const subscriptionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الاشتراكات الحالية");
  const packagesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الباقات");
  const newTeachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين"); // شيت المعلمين في الملف الجديد

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(120000); // زيادة مدة القفل لعملية النقل (2 دقيقة)

    // التحقق من وجود الشيتات الضرورية
    if (!oldStudentInfoSheet) throw new Error("شيت 'بيانات الطلبة' القديم غير موجود داخل الملف القديم. يرجى التأكد من اسمه الصحيح في الملف القديم.");
    if (!oldStudentScheduleSheet) throw new Error("شيت 'مواعيد الطلبة' القديم غير موجود داخل الملف القديم. يرجى التأكد من اسمه الصحيح.");
    if (!oldTeachersSheet) throw new Error("شيت 'المعلمين' القديم غير موجود داخل الملف القديم. يرجى التأكد من اسمه الصحيح.");
    if (!studentsSheet) throw new Error("شيت 'الطلاب' الجديد غير موجود.");
    if (!subscriptionsSheet) throw new Error("شيت 'الاشتراكات الحالية' الجديد غير موجود.");
    if (!packagesSheet) throw new Error("شيت 'الباقات' الجديد غير موجود.");
    if (!newTeachersSheet) throw new Error("شيت 'المعلمين' الجديد غير موجود.");

    const oldStudentInfoData = oldStudentInfoSheet.getDataRange().getValues();
    const oldStudentScheduleData = oldStudentScheduleSheet.getDataRange().getValues();
    const oldTeachersData = oldTeachersSheet.getDataRange().getValues();

    if (oldStudentInfoData.length < 2) {
      return { success: "لا توجد بيانات طلاب في شيت 'بيانات الطلبة' القديم لنقلها." };
    }

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    let migratedCount = 0;
    let failedCount = 0;
    const errors = [];
    const warnings = [];

    // 1. بناء خرائط من البيانات القديمة لسهولة البحث:
    // oldStudentInfoMap: key: Old Student ID (from بيانات الطلبة), value: { name, age, phone, subType, pkgName, attendedSessions }
    const oldStudentInfoMap = new Map();
    for (let i = 1; i < oldStudentInfoData.length; i++) {
        const row = oldStudentInfoData[i];
        const oldStudentId = String(row[0] || '').trim(); // العمود A
        if (oldStudentId) {
            oldStudentInfoMap.set(oldStudentId, {
                name: String(row[1] || '').trim(), // العمود B
                age: row[2],                        // العمود C
                phone: String(row[3] || '').trim(), // العمود D
                subscriptionType: String(row[6] || '').trim(), // العمود G
                packageName: String(row[7] || '').trim(),      // العمود H
                attendedSessions: row[8] || 0          // العمود I
            });
        }
    }
    Logger.log(`تم جمع ${oldStudentInfoMap.size} سجل طالب من 'بيانات الطلبة' القديم.`);


    // oldStudentScheduleMap: key: Old Student ID (from مواعيد الطلبة), value: { teacherName, startDate }
    const oldStudentScheduleMap = new Map();
    for (let i = 1; i < oldStudentScheduleData.length; i++) {
        const row = oldStudentScheduleData[i];
        const oldStudentId = String(row[0] || '').trim(); // العمود A
        if (oldStudentId) {
            oldStudentScheduleMap.set(oldStudentId, {
                teacherName: String(row[4] || '').trim(), // العمود E
                startDate: row[11] // العمود L
            });
        }
    }
    Logger.log(`تم جمع ${oldStudentScheduleMap.size} سجل مواعيد من 'مواعيد الطلبة' القديم.`);


    // oldTeacherNameToIdMap: key: Teacher Name (from المعلمين القديم), value: Teacher ID (from المعلمين القديم)
    const oldTeacherNameToIdMap = new Map();
    for (let i = 1; i < oldTeachersData.length; i++) {
        const teacherName = String(oldTeachersData[i][1] || '').trim(); // العمود B: اسم المعلم في الشيت القديم
        const teacherId = String(oldTeachersData[i][0] || '').trim(); // العمود A: معرف المعلم في الشيت القديم
        if (teacherName && teacherId) {
            oldTeacherNameToIdMap.set(teacherName, teacherId);
        }
    }
    Logger.log(`تم جمع ${oldTeacherNameToIdMap.size} معلم من 'المعلمين' القديم.`);


    // newTeacherIdToNameMap: key: Teacher ID (from المعلمين الجديد), value: Teacher Name (from المعلمين الجديد)
    const newTeacherIdToNameMap = new Map(); // هذه الخريطة ليست ضرورية للميجريشن ولكن مفيدة للمراجعة
    const newTeachersData = newTeachersSheet.getDataRange().getValues();
    for (let i = 1; i < newTeachersData.length; i++) {
        const teacherId = String(newTeachersData[i][0] || '').trim();
        const teacherName = String(newTeachersData[i][1] || '').trim();
        if (teacherId && teacherName) {
            newTeacherIdToNameMap.set(teacherId, teacherName);
        }
    }


    // packagesMap: key: packageName (from الباقات الجديد), value: { details }
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
                'نوع الباقة': String(packagesData[i][4] || '').trim(),
                'عدد الحصص الأسبوعية': packagesData[i][5] || 0
            });
        }
    }
    // خرائط التحويل لأسماء الباقات إذا تغيرت (احتفظ بها كما هي أو عدّلها لتناسب احتياجاتك)
    const oldToNewPackageNameMap = new Map([
        ["نص ساعة / 4 حصص", "4 حلقات / 30 دقيقة"],
        ["نصف ساعة / 8 حصص", "8 حلقات / 30 دقيقة"],
        ["ساعة / 4 حصص", "4 حلقات / 60 دقيقة"],
        ["ساعة / 8 حصص", "8 حلقات / 60 دقيقة"]
        // أضف أي تحويلات أخرى هنا
    ]);

    // 2. المرور على كل طالب في oldStudentInfoMap (لأنها المصدر الرئيسي للطلاب)
    for (const [oldStudentId, studentInfo] of oldStudentInfoMap.entries()) {
        try {
            const oldStudentName = studentInfo.name;
            const oldAge = parseFloat(studentInfo.age) || 0;
            const oldPhone = studentInfo.phone;
            const oldSubscriptionType = studentInfo.subscriptionType;
            let oldPackageName = studentInfo.packageName; // يمكن تعديلها
            const oldAttendedSessions = studentInfo.attendedSessions;

            // جلب معلومات المعلم وتاريخ بداية الاشتراك من 'مواعيد الطلبة' القديم
            const studentScheduleInfo = oldStudentScheduleMap.get(oldStudentId);
            let teacherNameFromSchedule = "";
            let startDateFromSchedule = ""; // ستحتوي على تاريخ بداية الاشتراك كـ Date object أو قيمة فارغة

            if (studentScheduleInfo) {
                teacherNameFromSchedule = studentScheduleInfo.teacherName;
                startDateFromSchedule = studentScheduleInfo.startDate;
            } else {
                warnings.push(`تحذير: لم يتم العثور على سجل مواعيد للطالب القديم ID: ${oldStudentId} (${oldStudentName}). سيتم تجاهل المعلم وتاريخ بداية الاشتراك.`);
            }

            // جلب Teacher ID من شيت المعلمين القديم
            const teacherId = oldTeacherNameToIdMap.get(teacherNameFromSchedule);
            if (!teacherId) {
                errors.push(`خطأ: لم يتم العثور على Teacher ID للمعلم '${teacherNameFromSchedule}' (الطالب: ${oldStudentName}, ID: ${oldStudentId}). سيتم تخطي هذا الطالب.`);
                failedCount++;
                continue;
            }

            // تحويل اسم الباقة القديم للجديد إذا وجد
            const newPackageName = oldToNewPackageNameMap.get(oldPackageName) || oldPackageName;

            // التحقق من وجود الباقة في الشيت الجديد
            const packageDetails = packagesMap.get(newPackageName);
            if (!packageDetails) {
                warnings.push(`تحذير: باقة '${newPackageName}' (محولة من '${oldPackageName}') للطالب ${oldStudentName} (ID: ${oldStudentId}). غير موجودة في شيت الباقات الجديد. سيتم إنشاء الاشتراك بقيم افتراضية للمبلغ.`);
            }

            // 3. إضافة الطالب إلى شيت "الطلاب" الجديد
            // لا ننشئ معرف طالب جديد، بل نستخدم المعرف القديم مباشرة
            const newStudentId = oldStudentId; // <--- هذا هو التعديل

            // تاريخ التسجيل في النظام الجديد سيكون تاريخ بداية الاشتراك الفعلي
            // إذا كان startDateFromSchedule موجودًا وصالحًا، استخدمه. وإلا، استخدم تاريخ اليوم.
            const registrationDate = startDateFromSchedule ? Utilities.formatDate(new Date(startDateFromSchedule), Session.getScriptTimeZone(), "yyyy-MM-dd") : today;

            let studentBasicStatus = "مشترك"; // افتراضي
            // يمكنك تحديد الحالة الأساسية للطالب هنا بناءً على منطق معين
            // مثلاً: إذا كان عدد الحصص الحاضرة صفر أو كان "نوع الاشتراك" تجريبي
            if (oldAttendedSessions === 0 && oldSubscriptionType === "تجريبي") {
                studentBasicStatus = "تجريبي";
            } else if (oldAttendedSessions === 0 && (oldPackageName.includes("لم يشترك") || oldPackageName.includes("معلق"))) {
                 studentBasicStatus = "معلق"; // مثال
            }

            studentsSheet.appendRow([
              newStudentId,           // A: Student ID
              oldStudentName,         // B: اسم الطالب
              oldAge,                 // C: السن
              oldPhone,               // D: رقم الهاتف (ولي الأمر)
              "",                     // E: رقم هاتف الطالب (إن وجد) - غير موجود في القديم
              "",                     // F: البلد - غير موجود في القديم
              registrationDate,       // G: تاريخ التسجيل
              studentBasicStatus,     // H: الحالة الأساسية للطالب
              ""                      // I: ملاحظات
            ]);
            Logger.log(`تم نقل الطالب ${oldStudentName} (ID القديم: ${oldStudentId}) إلى شيت 'الطلاب' الجديد بـ ID: ${newStudentId}.`);

            // 4. إنشاء اشتراك في شيت "الاشتراكات الحالية" الجديد
            const newSubscriptionId = generateUniqueSubscriptionId(subscriptionsSheet);

            let subscriptionAmount = packageDetails ? packageDetails['السعر'] || 0 : 0;
            let paidAmount = subscriptionAmount; // نفترض أنه مدفوع بالكامل عند النقل
            let remainingAmount = 0;
            let renewalStatus = "تم التجديد"; // نفترض أنه نشط عند النقل

            if (studentBasicStatus === "تجريبي") {
                paidAmount = 0;
                renewalStatus = "تجريبي";
                subscriptionAmount = 0; // الباقات التجريبية عادةً مجانية
            }


            // تاريخ بداية الاشتراك الفعلي في شيت الاشتراكات
            const actualSubscriptionStartDate = startDateFromSchedule ? Utilities.formatDate(new Date(startDateFromSchedule), Session.getScriptTimeZone(), "yyyy-MM-dd") : today;

            // حساب تاريخ نهاية الاشتراك المتوقع (الاعتماد على تاريخ بداية الاشتراك الفعلي)
            let endDate = "";
            if (packageDetails && packageDetails['نوع الباقة']) {
                if (packageDetails['نوع الباقة'] === "شهري") {
                    const startDateObj = new Date(actualSubscriptionStartDate); // استخدام التاريخ الفعلي هنا
                    startDateObj.setMonth(startDateObj.getMonth() + 1);
                    endDate = Utilities.formatDate(startDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
                } else if (packageDetails['نوع الباقة'] === "نصف سنوي") {
                    const startDateObj = new Date(actualSubscriptionStartDate); // استخدام التاريخ الفعلي هنا
                    startDateObj.setMonth(startDateObj.getMonth() + 6);
                    endDate = Utilities.formatDate(startDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
                } else if (packageDetails['نوع الباقة'] === "سنوي") {
                    const startDateObj = new Date(actualSubscriptionStartDate); // استخدام التاريخ الفعلي هنا
                    startDateObj.setFullYear(startDateObj.getFullYear() + 1);
                    endDate = Utilities.formatDate(startDateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
                }
            }


            subscriptionsSheet.appendRow([
              newSubscriptionId,           // A: Subscription ID
              newStudentId,                // B: Student ID
              newPackageName,              // C: اسم الباقة
              teacherId,                   // D: Teacher ID (المعلم الجديد)
              actualSubscriptionStartDate, // E: تاريخ بداية الاشتراك (من القديم أو تاريخ النقل)
              endDate,                     // F: تاريخ نهاية الاشتراك المتوقع
              oldAttendedSessions,         // G: عدد الحصص الحاضرة (من القديم)
              renewalStatus,               // H: الحالة التفصيلية للتجديد
              actualSubscriptionStartDate, // I: تاريخ آخر تجديد (نفس تاريخ البداية أو تاريخ النقل)
              subscriptionAmount,          // J: مبلغ الاشتراك الكلي
              paidAmount,                  // K: المبلغ المدفوع حتى الآن
              remainingAmount,             // L: المبلغ المتبقي
              oldSubscriptionType,         // M: نوع الاشتراك (من القديم)
              "تم النقل من النظام القديم" // N: ملاحظات خاصة بالاشتراك
            ]);
            Logger.log(`تم إنشاء اشتراك (${newSubscriptionId}) للطالب ${newStudentId} في شيت 'الاشتراكات الحالية'.`);

            migratedCount++;

        } catch (e) {
            errors.push(`خطأ في معالجة الطالب القديم ID: ${oldStudentId} (${studentInfo ? studentInfo.name : 'N/A'}): ${e.message}`);
            failedCount++;
            Logger.log(`فشل نقل الطالب ID ${oldStudentId}: ${e.message}`);
        }
    }

    const summary = `تمت عملية النقل بنجاح. تم نقل ${migratedCount} طالب. فشل نقل ${failedCount} طالب.`;
    Logger.log(summary);
    if (errors.length > 0) {
        Logger.log("تفاصيل الأخطاء:");
        errors.forEach(err => Logger.log(err));
    }
    if (warnings.length > 0) {
        Logger.log("تفاصيل التحذيرات:");
        warnings.forEach(warn => Logger.log(warn));
    }
    return { success: summary, errors: errors, warnings: warnings };

  } catch (e) {
    Logger.log("خطأ عام في عملية النقل: " + e.message);
    return { error: `فشل عام في عملية النقل: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}



/**
 * دالة لنقل سجلات الدفعات التاريخية من الملف القديم (شيت "السجل التاريخي")
 * إلى شيت "سجل الدفعات" في الملف الجديد.
 *
 * ستجلب البيانات من:
 * الملف القديم، شيت "السجل التاريخي": معرف الطالب (A), تاريخ الدفعة (F), المبلغ (G).
 *
 * ثم تضيفها إلى:
 * شيت "سجل الدفعات" الجديد.
 *
 * @returns {Object} كائن يحتوي على رسالة نجاح أو خطأ وملخص.
 */
function migratePaymentRecordsOnly() {
  // === قم بتغيير هذا الـ ID بالـ ID الخاص بملف Google Sheets القديم ===
  const OLD_SPREADSHEET_ID = "1XWRFKp-LM7detp42X4bqjhVgkaJT00of6FWvtx8sYL0"; // <--- تأكد من تحديث هذا الـ ID
  // ====================================================================

  let oldStudentsSpreadsheet; // اسمته oldStudentsSpreadsheet لأنه نفس الملف القديم بتاع الطلاب
  try {
    oldStudentsSpreadsheet = SpreadsheetApp.openById(OLD_SPREADSHEET_ID);
  } catch (e) {
    return { error: `فشل فتح ملف Google Sheets القديم بالـ ID المقدم: ${OLD_SPREADSHEET_ID}. الخطأ: ${e.message}` };
  }

  const oldHistoricalLogSheet = oldStudentsSpreadsheet.getSheetByName("السجل التاريخي"); // العمود A: معرف الطالب, F: تاريخ, G: المبلغ

  const paymentRecordsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("سجل الدفعات");
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("الطلاب"); // للتحقق من وجود الطالب في النظام الجديد

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(120000); // زيادة مدة القفل لعملية النقل (2 دقيقة)

    // التحقق من وجود الشيتات الضرورية
    if (!oldHistoricalLogSheet) throw new Error("شيت 'السجل التاريخي' القديم غير موجود داخل الملف القديم. يرجى التأكد من اسمه الصحيح.");
    if (!paymentRecordsSheet) throw new Error("شيت 'سجل الدفعات' الجديد غير موجود.");
    if (!studentsSheet) throw new Error("شيت 'الطلاب' الجديد غير موجود (لربط الدفعات بالطلاب الحاليين).");

    const oldHistoricalLogData = oldHistoricalLogSheet.getDataRange().getValues();

    if (oldHistoricalLogData.length < 2) {
      return { success: "لا توجد بيانات دفعات في شيت 'السجل التاريخي' القديم لنقلها." };
    }

    let migratedCount = 0;
    let failedCount = 0;
    const errors = [];
    const warnings = [];

    // بناء خريطة لـ Student IDs الموجودة في النظام الجديد لنتأكد أن الدفعة تُربط بطالب موجود.
    const newStudentIds = new Set();
    const newStudentsData = studentsSheet.getDataRange().getValues();
    for (let i = 1; i < newStudentsData.length; i++) {
        newStudentIds.add(String(newStudentsData[i][0] || '').trim()); // العمود A هو Student ID
    }
    Logger.log(`تم جمع ${newStudentIds.size} معرف طالب من شيت 'الطلاب' الجديد.`);

    // المرور على كل صف في شيت "السجل التاريخي" القديم
    for (let i = 1; i < oldHistoricalLogData.length; i++) {
      const row = oldHistoricalLogData[i];
      try {
        const oldStudentId = String(row[0] || '').trim(); // العمود A: معرف الطالب
        const paymentDateText = String(row[17] || '').trim(); // العمود F: تاريخ الدفعة (كنص)
        const paymentAmount = parseFloat(row[18]) || 0; // العمود G: المبلغ

        // تخطي الصفوف الفارغة أو اللي مفيهاش معرف طالب أو مبلغ
        if (!oldStudentId || paymentAmount <= 0) {
            warnings.push(`تحذير: تم تخطي الصف ${i + 1} في 'السجل التاريخي' لوجود بيانات ناقصة (ID: ${oldStudentId}, Amount: ${paymentAmount}).`);
            continue;
        }

        // التحقق من أن الطالب موجود في النظام الجديد قبل إضافة الدفعة
        if (!newStudentIds.has(oldStudentId)) {
            errors.push(`خطأ: الطالب بمعرف ${oldStudentId} غير موجود في شيت 'الطلاب' الجديد. لم يتم نقل الدفعة في الصف ${i + 1}.`);
            failedCount++;
            continue;
        }

        // إنشاء Payment ID جديد
        const newPaymentId = generateUniquePaymentId(paymentRecordsSheet);

        // ربط Subscription ID: سنتركه فارغًا كما اتفقنا لعدم وجود بيانات موثوقة في القديم.
        const subscriptionId = ''; // يتم تركه فارغاً

        // طريقة الدفع
        const paymentMethod = "فودافون كاش"; // ثابت كما طلب

        // ملاحظات (يمكن إضافة ملاحظة توضح أنها دفعة من الهجرة)
        const paymentNotes = "دفعة من السجل التاريخي القديم";

        // إضافة الدفعة إلى شيت "سجل الدفعات" الجديد
        // الأعمدة الجديدة: Payment ID (0) - Subscription ID (1) - Student ID (2) - تاريخ الدفع (3) - المبلغ المدفوع (4) - طريقة الدفع (5) - ملاحظات (6)
        paymentRecordsSheet.appendRow([
          newPaymentId,      // Payment ID (A)
          subscriptionId,    // Subscription ID (B)
          oldStudentId,      // Student ID (C)
          paymentDateText,   // تاريخ الدفع (D) - كنص
          paymentAmount,     // المبلغ المدفوع (E)
          paymentMethod,     // طريقة الدفع (F)
          paymentNotes       // ملاحظات (G)
        ]);
        Logger.log(`تم نقل دفعة لـ ${oldStudentId} بمبلغ ${paymentAmount} وتاريخ ${paymentDateText}.`);

        migratedCount++;

      } catch (e) {
        errors.push(`خطأ في معالجة سجل الدفعة في الصف ${i + 1} (الطالب ID: ${row[0] || 'N/A'}): ${e.message}`);
        failedCount++;
        Logger.log(`فشل نقل دفعة في الصف ${i + 1}: ${e.message}`);
      }
    }

    const summary = `تمت عملية نقل سجلات الدفعات بنجاح. تم نقل ${migratedCount} دفعة. فشل نقل ${failedCount} دفعة.`;
    Logger.log(summary);
    if (errors.length > 0) {
        Logger.log("تفاصيل الأخطاء:");
        errors.forEach(err => Logger.log(err));
    }
    if (warnings.length > 0) {
        Logger.log("تفاصيل التحذيرات:");
        warnings.forEach(warn => Logger.log(warn));
    }
    return { success: summary, errors: errors, warnings: warnings };

  } catch (e) {
    Logger.log("خطأ عام في عملية نقل سجلات الدفعات: " + e.message);
    return { error: `فشل عام في عملية نقل سجلات الدفعات: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


/**
 * تسجل حلقة احتياطية لمرة واحدة لمعلم بديل في شيت "المواعيد المتاحة للمعلمين"
 * من خلال وضع القيمة "p studentId" في الخلية المناسبة.
 *
 * @param {string} studentId - رقم الطالب
 * @param {string} studentName - اسم الطالب
 * @param {string} backupTeacherId - معرّف المعلم البديل
 * @param {string} day - اليوم (مثال: "الأحد")
 * @param {string} timeSlot - ميعاد الحصة (مثال: "5:00 مساءً")
 * @returns {Object} - رسالة نجاح أو خطأ
 */
function assignBackupSession(studentId, studentName, backupTeacherId, day, timeSlot) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("المواعيد المتاحة للمعلمين");
    if (!sheet) throw new Error("شيت المواعيد غير موجود");

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const dayColIndex = headers.indexOf(day);
    if (dayColIndex === -1) throw new Error(`لم يتم العثور على العمود المناسب لليوم: ${day}`);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const teacherIdInRow = String(row[0] || "").trim(); // العمود A
      const timeSlotInRow = String(row[1] || "").trim();  // العمود B

      if (teacherIdInRow === backupTeacherId && timeSlotInRow === timeSlot) {
        const cellValue = String(row[dayColIndex] || "").trim();

        if (cellValue && !cellValue.startsWith("p")) {
          return { error: "هذا الميعاد غير متاح، أو محجوز بالفعل." };
        }

        // 🟢 تسجيل الحلقة الاحتياطية
        sheet.getRange(i + 1, dayColIndex + 1).setValue("p " + studentId);
        return { success: `تم تسجيل الحلقة الاحتياطية بنجاح مع المعلم في ${day} الساعة ${timeSlot}` };
      }
    }

    return { error: "لم يتم العثور على صف مناسب لهذا المعلم والميعاد." };

  } catch (err) {
    return { error: "حدث خطأ أثناء تسجيل الحلقة الاحتياطية: " + err.message };
  }
}





/**
 * تُستخدم لتحميل بيانات نموذج الحلقة الاحتياطية في واجهة المشرف.
 * تُعيد قائمة الطلاب و المواعيد المتاحة.
 */
function getBackupFormOptions() {
  const ss = SpreadsheetApp.openById("11jAQXDKzwV--h7sNkvESm0dvNcRNk36Z3IenOtLIdsY");
  const studentsSheet = ss.getSheetByName("الطلاب");
  const slotsSheet = ss.getSheetByName("المواعيد المتاحة للمعلمين");

  const studentsData = studentsSheet.getDataRange().getValues();
  const students = studentsData.slice(1).map(row => ({
    id: row[0]?.toString().trim(),
    name: row[1]?.toString().trim()
  }));

  const headers = slotsSheet.getDataRange().getValues()[0];
  const timeSlots = headers.slice(2).filter(Boolean); // الأعمدة بعد العمودين الأولين

  return { students, timeSlots };
}



function getAllStudentsForReserveFeature(teacherId) {
  const allStudents = getAllStudentsForTeacher(teacherId);
  return Array.isArray(allStudents) ? allStudents : [];
}


/**
 * تجلب بيانات جميع الطلاب (المشتركين والتجريبيين) المرتبطين بمعلم محدد.
 * تُستخدم لصفحة "طلابي ومواعيدهم" في واجهة المعلم.
 *
 * @param {string} teacherId - Teacher ID للمعلم المراد جلب طلابه.
 * @returns {Array<Object>} مصفوفة من كائنات الطلاب الموحدة.
 * أو {Object} كائن خطأ.
 */
function getAllStudentsForTeacher(teacherId) {
  const supervisorSpreadsheet = SpreadsheetApp.openById("11jAQXDKzwV--h7sNkvESm0dvNcRNk36Z3IenOtLIdsY");
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

  // جلب بيانات الطلاب (ID -> Name) من كلا الشيتين في ملف المشرف
  const studentIdToNameMap = new Map();
  studentsSheet.getDataRange().getValues().forEach(row => {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const phone = String(row[3] || '').trim(); // جلب رقم الهاتف
    const basicStatus = String(row[7] || '').trim(); // جلب الحالة الأساسية
    if (id) studentIdToNameMap.set(id, { name: name, phone: phone, basicStatus: basicStatus });
  });
  trialStudentsSheet.getDataRange().getValues().forEach(row => {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const phone = String(row[3] || '').trim(); // جلب رقم الهاتف
    const basicStatus = String(row[10] || '').trim(); // الحالة من عمود Status في شيت التجريبيين
    if (id) studentIdToNameMap.set(id, { name: name, phone: phone, basicStatus: basicStatus });
  });

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
  studentIdToNameMap.forEach((studentDetails, studentID) => {
    if (studentID.startsWith("STD")) { // فقط الطلاب المشتركين (وليس التجريبيين)
      const subscriptionDetails = subscriptionsMap.get(studentID);
      if (subscriptionDetails) { // إذا كان الطالب مشتركًا لهذا المعلم
        const studentInfo = {
          studentID: studentID,
          name: studentDetails.name,
          age: null, // لا يتم جلبه من هنا، يمكن جلبه من شيت الطلاب إذا لزم الأمر
          phone: studentDetails.phone,
          basicStatus: studentDetails.basicStatus, // الحالة الأساسية
          packageName: subscriptionDetails.packageName,
          renewalStatus: subscriptionDetails.renewalStatus,
        };

        const bookedSlots = studentBookedSlotsMap.get(studentID) || [];
        
        // التعديل: إضافة مصفوفة بجميع المواعيد المحجوزة
        studentInfo.allBookedScheduleSlots = bookedSlots.map(slot => ({
            day: slot.day,
            time: slot.timeSlotHeader
        })).sort((a,b) => {
            const daysOrder = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
            const dayAIndex = daysOrder.indexOf(a.day);
            const dayBIndex = daysOrder.indexOf(b.day);
            if (dayAIndex !== dayBIndex) return dayAIndex - dayBIndex;
            return getTimeInMinutes(a.time) - getTimeInMinutes(b.time);
        });
        
        allTeacherStudents.push(studentInfo);
      }
    } else if (studentID.startsWith("TRL")) { // الطلاب التجريبيون
        const trialStudentRawData = trialStudentsSheet.getDataRange().getValues();
        const trialRow = trialStudentRawData.find(r => String(r[0] || '').trim() === studentID);
        if (trialRow && String(trialRow[4] || '').trim() === teacherId) { // التأكد أنه لهذا المعلم
          const trialStudentInfo = {
              studentID: trialID, // هنا هو الـ Trial ID
              name: studentDetails.name,
              age: trialRow[2], // السن من شيت الطلاب التجريبيين
              phone: studentDetails.phone,
              basicStatus: studentDetails.basicStatus,
              packageName: 'تجريبي',
              renewalStatus: 'تجريبي',
          };
          trialStudentInfo.allBookedScheduleSlots = [];
          if (String(trialRow[6] || '').trim() && String(trialRow[7] || '').trim()) {
              trialStudentInfo.allBookedScheduleSlots.push({
                  day: String(trialRow[6] || '').trim(), // اليوم الأول
                  time: String(trialRow[7] || '').trim() // الميعاد الأول
              });
          }
          allTeacherStudents.push(trialStudentInfo);
        }
    }
  });

  // --- معالجة الحلقات الاحتياطية لهذا المعلم ---
const backupSessionsSheet = supervisorSpreadsheet.getSheetByName("الحلقات الاحتياطية");
if (backupSessionsSheet) {
  const backupData = backupSessionsSheet.getDataRange().getValues();
  const todayName = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"][new Date().getDay()];
  
  backupData.forEach((row, index) => {
    if (index === 0) return; // تجاهل العناوين
    const [studentID, day, time, subject, mainTeacherId, backupTeacherId, status] = row.map(v => String(v).trim());
    if (backupTeacherId === teacherId && day === todayName && !status) {
      const studentDetails = studentIdToNameMap.get(studentID);
      if (studentDetails) {
        const studentInfo = {
          studentID: studentID,
          name: studentDetails.name,
          age: null,
          phone: studentDetails.phone,
          basicStatus: studentDetails.basicStatus || '',
          packageName: 'احتياطي',
          renewalStatus: 'احتياطي',
          isBackup: true
        };
        studentInfo.allBookedScheduleSlots = [{ day, time }];
        allTeacherStudents.push(studentInfo);
      }
    }
  });
}


  Logger.log(`تم جلب ${allTeacherStudents.length} طالب للمعلم ID ${teacherId}.`);
  return allTeacherStudents;
}


/**
 * تجلب قائمة بجميع المعلمين من شيت "المعلمين" مع المعرف والاسم.
 * @returns {Array<Object>} قائمة تحتوي على كل معلم ككائن {id, name}
 */
function getAllTeachersWithIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("المعلمين");
  if (!sheet) {
    Logger.log("❌ خطأ: لم يتم العثور على شيت 'المعلمين'.");
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const teachers = [];

  for (let i = 1; i < data.length; i++) { // تخطي العنوان
    const id = String(data[i][0] || '').trim();   // العمود A = ID
    const name = String(data[i][1] || '').trim(); // العمود B = الاسم

    if (id && name) {
      teachers.push({ id, name });
    }
  }

  Logger.log("✅ قائمة المعلمين بالمعرفات: " + JSON.stringify(teachers));
  return teachers;
}
