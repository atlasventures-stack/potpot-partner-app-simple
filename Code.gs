// ============================================
  // POTPOT BOOKING SYSTEM - Code.gs
  // ============================================

  const TABS = {
    GARDENER_ZONES: 'GardenerZones',
    AVAILABILITY: 'Availability',
    BOOKINGS: 'Bookings',
    SERVICE_REPORTS: 'ServiceReports',
    ADMINS: 'Admins'
  };

  // ============================================
  // WATI.IO CONFIGURATION
  // ============================================

  const WATI_API_ENDPOINT = 'https://live-mt-server.wati.io/1071439';
  const WATI_ACCESS_TOKEN = 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1bmlxdWVfbmFtZSI6ImF2aXNoajE2QGdtYWlsLmNvbSIsIm5hbWVpZCI6ImF2aXNoajE2QGdtYWlsLmNvbSIsImVtYWlsIjoiYXZpc2hqMTZAZ21haWwuY29tIiwiYXV0aF90aW1lIjoiMDEvMDMvMjAyNiAxNDowNzo0NyIsInRlbmFudF9pZCI6IjEwNzE0MzkiLCJkYl9uYW1lIjoibXQtcHJvZC1UZW5hbnRzIiwiaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS93cy8yMDA4LzA2L2lkZW50aXR5L2NsYWltcy9yb2xlIjoiQURNSU5JU1RSQVRPUiIsImV4cCI6MjUzNDAyMzAwODAwLCJpc3MiOiJDbGFyZV9BSSIsImF1ZCI6IkNsYXJlX0FJIn0.QhXKWW4bUqDRviuFuiwFTaTD2r5zQYApVDo6fOFaTQg';

  // ============================================
  // RAZORPAY CONFIGURATION
  // ============================================

  const RAZORPAY_KEY_ID = 'rzp_live_S1O9qIcBWYvsPI';
  const RAZORPAY_KEY_SECRET = 'CJ62JCmvMUjUaL2YpciR8jh5';

  // ============================================
  // UNIVERSAL DATE PARSER - Handles all formats
  // ============================================

  function parseDate(dateValue) {
    if (!dateValue || dateValue === 'Not yet' || dateValue === 'not yet') return null;

    // Handle Date objects from Google Sheets (instanceof may fail across realms)
    if (dateValue instanceof Date ||
        (dateValue && typeof dateValue.getTime === 'function')) {
      if (isNaN(dateValue.getTime())) return null;
      return dateValue;
    }

    // Also check for Date-like objects from Sheets that might not pass instanceof
    if (dateValue && typeof dateValue === 'object' && dateValue.constructor &&
        dateValue.constructor.name === 'Date') {
      if (isNaN(dateValue.getTime())) return null;
      return dateValue;
    }

    const str = String(dateValue).trim();
    if (!str) return null;

    // DD/MM/YYYY or DD-MM-YYYY
    let match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = parseInt(match[2], 10) - 1;
      const year = parseInt(match[3], 10);
      return new Date(year, month, day);
    }

    // YYYY-MM-DD or YYYY/MM/DD
    match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (match) {
      const year = parseInt(match[1], 10);
      const month = parseInt(match[2], 10) - 1;
      const day = parseInt(match[3], 10);
      return new Date(year, month, day);
    }

    // Try native Date parsing as fallback
    const fallback = new Date(str);
    if (!isNaN(fallback.getTime())) return fallback;

    return null;
  }

  // ============================================
  // CUSTOM MENU - Shows when spreadsheet opens
  // ============================================

  function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🌱 PotPot')
      .addItem('📩 Send Reminder to Selected Row', 'sendReminderToSelectedRow')
      .addItem('📩 Send Reminders to All Tomorrow Bookings', 'sendRemindersToTomorrow')
      .addSeparator()
      .addItem('📩 Send 5-Day Follow-ups', 'sendFiveDayFollowups')
      .addToUi();
  }

  // ============================================
  // SEND REMINDER TO SELECTED ROW
  // ============================================

  function sendReminderToSelectedRow() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TABS.BOOKINGS);
    const ui = SpreadsheetApp.getUi();

    const selection = ss.getSelection();
    const activeRange = selection.getActiveRange();

    if (!activeRange) {
      ui.alert('Error', 'Please select a row in the Bookings sheet first.', ui.ButtonSet.OK);
      return;
    }

    const row = activeRange.getRow();

    if (row < 2) {
      ui.alert('Error', 'Please select a data row (not the header).', ui.ButtonSet.OK);
      return;
    }

    const rowData = sheet.getRange(row, 1, 1, 15).getValues()[0];

    const customerName = rowData[1];
    const phone = rowData[2];
    const reachTime = rowData[11];

    if (!phone) {
      ui.alert('Error', 'No phone number found for this booking.', ui.ButtonSet.OK);
      return;
    }

    if (!reachTime) {
      ui.alert('Error', 'No ReachTime found for this booking. Please fill Column L first.', ui.ButtonSet.OK);
      return;
    }

    const response = ui.alert(
      'Send Reminder',
      `Send reminder to ${customerName}?\n\nPhone: ${phone}\nReach Time: ${reachTime}`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    const result = sendVisitReminder(phone, reachTime);

    if (result.result === true) {
      ui.alert('Success', `Reminder sent to ${customerName} (${phone})`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Failed to send reminder: ${JSON.stringify(result)}`, ui.ButtonSet.OK);
    }
  }

  // ============================================
  // SEND REMINDERS TO ALL TOMORROW'S BOOKINGS
  // ============================================

  function sendRemindersToTomorrow() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TABS.BOOKINGS);
    const ui = SpreadsheetApp.getUi();

    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    const data = sheet.getDataRange().getValues();
    const tomorrowBookings = [];

    for (let i = 1; i < data.length; i++) {
      const bookingDate = data[i][6];
      const bookingDateStr = parseDateToFormatted(bookingDate);

      if (bookingDateStr === tomorrowStr) {
        const reachTime = data[i][11];
        if (reachTime) {
          tomorrowBookings.push({
            row: i + 1,
            customerName: data[i][1],
            phone: data[i][2],
            reachTime: reachTime
          });
        }
      }
    }

    if (tomorrowBookings.length === 0) {
      ui.alert('No Bookings', 'No bookings with ReachTime found for tomorrow.', ui.ButtonSet.OK);
      return;
    }

    const names = tomorrowBookings.map(b => b.customerName).join(', ');
    const response = ui.alert(
      'Send Reminders',
      `Send reminders to ${tomorrowBookings.length} customer(s)?\n\n${names}`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    let successCount = 0;
    let failCount = 0;

    for (const booking of tomorrowBookings) {
      if (!booking.phone) {
        failCount++;
        continue;
      }

      const result = sendVisitReminder(booking.phone, booking.reachTime);

      if (result.result === true) {
        successCount++;
      } else {
        failCount++;
      }

      Utilities.sleep(500);
    }

    ui.alert('Done', `Reminders sent!\n\nSuccess: ${successCount}\nFailed: ${failCount}`, ui.ButtonSet.OK);
  }

  // ============================================
  // SEND 5-DAY FOLLOW-UP (Manual - From Menu)
  // ============================================

  function sendFiveDayFollowups() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CustomerData');
    const ui = SpreadsheetApp.getUi();

    if (!sheet) {
      ui.alert('Error', 'CustomerData sheet not found.', ui.ButtonSet.OK);
      return;
    }

    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const eligibleCustomers = [];

    for (let i = 1; i < data.length; i++) {
      const name = data[i][1];
      const phone = data[i][2];
      const lastServiceDate = data[i][14];

      if (!phone || !lastServiceDate) continue;

      const serviceDate = parseDate(lastServiceDate);
      if (!serviceDate) continue;

      serviceDate.setHours(0, 0, 0, 0);

      const diffTime = today.getTime() - serviceDate.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays === 5) {
        eligibleCustomers.push({
          row: i + 1,
          name: name || 'Customer',
          phone: phone,
          lastServiceDate: serviceDate,
          daysSince: diffDays
        });
      }
    }

    if (eligibleCustomers.length === 0) {
      ui.alert('No Customers', 'No customers found with last service 5 days ago.', ui.ButtonSet.OK);
      return;
    }

    const names = eligibleCustomers.map(c => `${c.name} (${c.daysSince} days)`).join('\n');
    const response = ui.alert(
      'Send 5-Day Follow-ups',
      `Send follow-up to ${eligibleCustomers.length} customer(s)?\n\n${names}`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    let successCount = 0;
    let failCount = 0;

    for (const customer of eligibleCustomers) {
      const result = sendPostServiceCheckup(customer.phone);

      if (result.result === true) {
        successCount++;
      } else {
        failCount++;
      }

      Utilities.sleep(500);
    }

    ui.alert('Done', `Follow-ups sent!\n\nSuccess: ${successCount}\nFailed: ${failCount}`, ui.ButtonSet.OK);
  }

  // ============================================
  // AUTO SEND 5-DAY FOLLOW-UPS (For Daily Trigger)
  // ============================================

  function autoSendFiveDayFollowups() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CustomerData');

    if (!sheet) {
      Logger.log('ERROR: CustomerData sheet not found');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let successCount = 0;
    let failCount = 0;
    const sentTo = [];

    for (let i = 1; i < data.length; i++) {
      const name = data[i][1];
      const phone = data[i][2];
      const lastServiceDate = data[i][14];

      if (!phone || !lastServiceDate) continue;

      const serviceDate = parseDate(lastServiceDate);
      if (!serviceDate) continue;

      serviceDate.setHours(0, 0, 0, 0);

      const diffTime = today.getTime() - serviceDate.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays === 5) {
        const result = sendPostServiceCheckup(phone);

        if (result.result === true) {
          successCount++;
          sentTo.push(name + ' (' + phone + ')');
          Logger.log('✅ Sent to: ' + name + ' (' + phone + ')');
        } else {
          failCount++;
          Logger.log('❌ Failed for: ' + name + ' (' + phone + ') - ' + JSON.stringify(result));
        }

        Utilities.sleep(500);
      }
    }

    Logger.log('=== DAILY FOLLOWUP COMPLETE ===');
    Logger.log('Date: ' + today.toDateString());
    Logger.log('Success: ' + successCount + ', Failed: ' + failCount);
    if (sentTo.length > 0) {
      Logger.log('Sent to: ' + sentTo.join(', '));
    }
  }

  // ============================================
  // WATI.IO - SEND VISIT REMINDER (Day Before)
  // ============================================

  function sendVisitReminder(phoneNumber, timeSlot) {
    const cleanPhone = String(phoneNumber).replace(/\D/g, '').slice(-10);

    const url = `${WATI_API_ENDPOINT}/api/v1/sendTemplateMessage?whatsappNumber=91${cleanPhone}`;

    const payload = {
      "template_name": "pre_service_com_time",
      "broadcast_name": "pre_service_reminder_" + Date.now(),
      "parameters": [
        { "name": "1", "value": String(timeSlot) }
      ]
    };

    const options = {
      method: 'POST',
      headers: {
        'Authorization': WATI_ACCESS_TOKEN,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      Logger.log('Wati Reminder Response: ' + response.getContentText());
      return JSON.parse(response.getContentText());
    } catch (error) {
      Logger.log('Wati Reminder Error: ' + error);
      return { error: error.message };
    }
  }

  // ============================================
  // WATI.IO - SEND POST SERVICE CHECKUP (5-Day)
  // ============================================

  function sendPostServiceCheckup(phoneNumber) {
    const cleanPhone = String(phoneNumber).replace(/\D/g, '').slice(-10);

    const url = `${WATI_API_ENDPOINT}/api/v1/sendTemplateMessage?whatsappNumber=91${cleanPhone}`;

    const payload = {
      "template_name": "post_service_checkup",
      "broadcast_name": "post_service_checkup_" + Date.now(),
      "parameters": []
    };

    const options = {
      method: 'POST',
      headers: {
        'Authorization': WATI_ACCESS_TOKEN,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      Logger.log('Wati Checkup Response: ' + response.getContentText());
      return JSON.parse(response.getContentText());
    } catch (error) {
      Logger.log('Wati Checkup Error: ' + error);
      return { error: error.message };
    }
  }

  // ============================================
  // WATI.IO - SEND NPS FORM (After Service)
  // ============================================

  function sendNPSForm(phoneNumber) {
    const cleanPhone = String(phoneNumber).replace(/\D/g, '').slice(-10);

    const url = `${WATI_API_ENDPOINT}/api/v1/sendTemplateMessage?whatsappNumber=91${cleanPhone}`;

    const payload = {
      "template_name": "nps_form",
      "broadcast_name": "nps_form_" + Date.now(),
      "parameters": []
    };

    const options = {
      method: 'POST',
      headers: {
        'Authorization': WATI_ACCESS_TOKEN,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      Logger.log('NPS Form Response: ' + response.getContentText());
      return JSON.parse(response.getContentText());
    } catch (error) {
      Logger.log('NPS Form Error: ' + error);
      return { error: error.message };
    }
  }

  // ============================================
  // RAZORPAY - CREATE PAYMENT LINK
  // ============================================

  function createPaymentLink(data) {
    try {
      const amount = parseInt(data.amount) * 100; // Convert to paise
      const customerName = data.customerName || 'Customer';
      const customerPhone = data.customerPhone ? String(data.customerPhone).replace(/\D/g, '').slice(-10) : '';
      const bookingID = data.bookingID;

      const payload = {
        amount: amount,
        currency: 'INR',
        accept_partial: false,
        description: 'PotPot Plant Care Service - Booking ' + bookingID,
        customer: {
          name: customerName,
          contact: customerPhone ? '+91' + customerPhone : undefined
        },
        notify: {
          sms: false,
          email: false
        },
        reminder_enable: false,
        notes: {
          booking_id: bookingID
        },
        expire_by: Math.floor(Date.now() / 1000) + (24 * 60 * 60) // Expires in 24 hours
      };

      // Remove undefined fields
      if (!payload.customer.contact) delete payload.customer.contact;

      const options = {
        method: 'POST',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(RAZORPAY_KEY_ID + ':' + RAZORPAY_KEY_SECRET),
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.razorpay.com/v1/payment_links', options);
      const result = JSON.parse(response.getContentText());

      if (result.id) {
        Logger.log('Payment link created: ' + result.short_url);
        return {
          success: true,
          paymentLinkId: result.id,
          paymentLink: result.short_url
        };
      } else {
        Logger.log('Razorpay error: ' + JSON.stringify(result));
        return { success: false, error: result.error ? result.error.description : 'Failed to create payment link' };
      }
    } catch (error) {
      Logger.log('createPaymentLink error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  }

  // ============================================
  // RAZORPAY - CHECK PAYMENT STATUS
  // ============================================

  function checkPaymentStatus(paymentLinkId) {
    try {
      const options = {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(RAZORPAY_KEY_ID + ':' + RAZORPAY_KEY_SECRET)
        },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.razorpay.com/v1/payment_links/' + paymentLinkId, options);
      const result = JSON.parse(response.getContentText());

      if (result.id) {
        // Payment link statuses: created, partially_paid, paid, cancelled, expired
        const isPaid = result.status === 'paid' || result.amount_paid >= result.amount;
        return {
          success: true,
          status: isPaid ? 'paid' : 'pending',
          amountPaid: result.amount_paid / 100,
          totalAmount: result.amount / 100
        };
      } else {
        return { success: false, error: 'Payment link not found' };
      }
    } catch (error) {
      Logger.log('checkPaymentStatus error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  }

  // ============================================
  // UPDATE PAYMENT STATUS IN SERVICE REPORT
  // ============================================

  function updatePaymentStatus(data) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const reportsSheet = ss.getSheetByName(TABS.SERVICE_REPORTS);

      if (!reportsSheet) {
        return { success: false, error: 'ServiceReports sheet not found' };
      }

      const reportsData = reportsSheet.getDataRange().getValues();

      // Find the row with matching bookingID
      for (let i = 1; i < reportsData.length; i++) {
        if (reportsData[i][1] === data.bookingID) {
          // Update PaymentStatus (Column O = index 14) and PaymentLinkId (Column P = index 15)
          reportsSheet.getRange(i + 1, 15).setValue(data.status || 'Paid');
          reportsSheet.getRange(i + 1, 16).setValue(data.paymentLinkId || '');

          Logger.log('Payment status updated for booking: ' + data.bookingID);
          return { success: true, message: 'Payment status updated' };
        }
      }

      return { success: false, error: 'Booking not found in reports' };
    } catch (error) {
      Logger.log('updatePaymentStatus error: ' + error.toString());
      return { success: false, error: error.toString() };
    }
  }

  // ============================================
  // API ENDPOINTS
  // ============================================

  function doGet(e) {
    try {
      const action = e.parameter.action;

      if (action === 'validatePincode') {
        const pincode = e.parameter.pincode;
        if (!pincode) {
          return jsonResponse({ valid: false, error: 'Pincode required' });
        }
        const isValid = validatePincode(pincode);
        return jsonResponse({ valid: isValid });
      }

      if (action === 'getGardenerJobs') {
        const phone = e.parameter.phone;
        if (!phone) {
          return jsonResponse({ success: false, error: 'Phone is required' });
        }
        const result = getGardenerJobs(phone);
        return jsonResponse(result);
      }

      if (action === 'getAdminReports') {
        const phone = e.parameter.phone;
        if (!phone) {
          return jsonResponse({ success: false, error: 'Phone is required' });
        }
        const result = getAdminReports(phone);
        return jsonResponse(result);
      }

      if (action === 'checkPaymentStatus') {
        const paymentLinkId = e.parameter.paymentLinkId;
        if (!paymentLinkId) {
          return jsonResponse({ success: false, error: 'Payment link ID required' });
        }
        const result = checkPaymentStatus(paymentLinkId);
        return jsonResponse(result);
      }

      const pincode = e.parameter.pincode;
      const lat = e.parameter.lat ? parseFloat(e.parameter.lat) : null;
      const lng = e.parameter.lng ? parseFloat(e.parameter.lng) : null;
      const plantCount = e.parameter.plantCount ? parseInt(e.parameter.plantCount) : 20; // Default to 20 plants (1hr service)

      if (!pincode && !lat) {
        return jsonResponse({ error: 'Pincode or coordinates required' });
      }

      const slots = getAvailableSlots(pincode, lat, lng, plantCount);
      return jsonResponse({ success: true, slots: slots });
    } catch (error) {
      return jsonResponse({ error: error.toString() });
    }
  }

  function doPost(e) {
    try {
      const data = JSON.parse(e.postData.contents);

      if (data.action === 'saveServiceReport') {
        const result = saveServiceReport(data);
        return jsonResponse(result);
      }

      if (data.action === 'createPaymentLink') {
        const result = createPaymentLink(data);
        return jsonResponse(result);
      }

      if (data.action === 'updatePaymentStatus') {
        const result = updatePaymentStatus(data);
        return jsonResponse(result);
      }

      const result = createBooking(data);
      return jsonResponse({ success: true, booking: result });
    } catch (error) {
      return jsonResponse({ error: error.toString() });
    }
  }

  function jsonResponse(data) {
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }

  // ============================================
  // VALIDATE PINCODE (Early check at Step 1)
  // ============================================

  function validatePincode(pincode) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const zonesSheet = ss.getSheetByName(TABS.GARDENER_ZONES);
    const zonesData = zonesSheet.getDataRange().getValues();

    const inputPincode = String(pincode).trim().replace('.0', '');

    for (let i = 1; i < zonesData.length; i++) {
      const rawPincodes = String(zonesData[i][0] || '');
      const pincodes = rawPincodes.split(',').map(p => String(p).trim().replace('.0', ''));

      if (pincodes.includes(inputPincode)) {
        return true;
      }
    }

    return false;
  }

  // ============================================
  // WATI.IO - SEND BOOKING CONFIRMATION
  // ============================================

  function sendBookingConfirmation(phoneNumber, date, timeSlot) {
    const cleanPhone = String(phoneNumber).replace(/\D/g, '').slice(-10);

    const url = `${WATI_API_ENDPOINT}/api/v1/sendTemplateMessage?whatsappNumber=91${cleanPhone}`;

    const payload = {
      "template_name": "booking_confirmation",
      "broadcast_name": "booking_confirm_" + Date.now(),
      "parameters": [
        { "name": "1", "value": date },
        { "name": "2", "value": timeSlot }
      ]
    };

    const options = {
      method: 'POST',
      headers: {
        'Authorization': WATI_ACCESS_TOKEN,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      Logger.log('Wati Response: ' + response.getContentText());
      return JSON.parse(response.getContentText());
    } catch (error) {
      Logger.log('Wati Error: ' + error);
      return { error: error.message };
    }
  }

  // ============================================
  // HELPER: Parse date string to dd/MM/yyyy
  // ============================================

  function parseDateToFormatted(dateValue) {
    const parsed = parseDate(dateValue);
    if (parsed) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return String(dateValue);
  }

  // ============================================
  // SLOT SYSTEM CONFIGURATION
  // ============================================

  const SLOT_CONFIG = {
    START_HOUR: 8,
    START_MIN: 30,
    END_HOUR: 18,        // Last slot at 6:00 PM
    END_MIN: 0,
    SLOT_INTERVAL: 30,   // 30 min slots
    LUNCH_START: 13,     // 1:00 PM
    LUNCH_END: 14,       // 2:00 PM
    HARD_CUTOFF: 18.5,   // 6:30 PM - all work must finish
    TRAVEL_BUFFER: 30    // 30 min travel after each booking
  };

  // ============================================
  // HELPER: Get service duration from plant count
  // ============================================

  function getServiceDuration(plantCount) {
    // Handle empty/null values
    if (!plantCount) return 90;  // Default 1.5 hours for old bookings without plantCount

    let count = 0;
    const plantStr = String(plantCount);

    // Handle string ranges like "0-20", "20-35", "35-50", "50+"
    if (plantStr.includes('-')) {
      // Extract upper bound: "0-20" -> 20, "20-35" -> 35
      const parts = plantStr.split('-');
      count = parseInt(parts[1]) || parseInt(parts[0]) || 0;
    } else if (plantStr.includes('+')) {
      // Handle "50+" format -> 51
      count = parseInt(plantStr) + 1 || 51;
    } else {
      // Plain number
      count = parseInt(plantStr) || 0;
    }

    if (count === 0) return 90;       // Default 1.5 hours for old bookings without plantCount
    if (count <= 20) return 60;       // 1 hour
    if (count <= 35) return 90;       // 1.5 hours
    if (count <= 50) return 120;      // 2 hours
    return 180;                        // 3 hours (51-150 plants)
  }

  // ============================================
  // HELPER: Time string to minutes from midnight
  // ============================================

  function timeToMinutes(timeStr) {
    if (!timeStr) return 0;
    const str = String(timeStr);

    // Handle "8:30 AM" or "8:30:00 AM" format (12-hour with AM/PM, optional seconds)
    // Use ^ anchor to match from start and avoid matching wrong part of string
    const match12hr = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)/i);
    if (match12hr) {
      let hours = parseInt(match12hr[1]);
      const mins = parseInt(match12hr[2]);
      const period = match12hr[3].toUpperCase();

      if (period === 'PM' && hours !== 12) hours += 12;
      if (period === 'AM' && hours === 12) hours = 0;

      return hours * 60 + mins;
    }

    // Handle "14:00" or "14:00:00" format (24-hour without AM/PM)
    const match24hr = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (match24hr) {
      const hours = parseInt(match24hr[1]);
      const mins = parseInt(match24hr[2]);
      return hours * 60 + mins;
    }

    // Fallback: try to find any time pattern in the string (for long date strings)
    const matchAny = str.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)/i);
    if (matchAny) {
      let hours = parseInt(matchAny[1]);
      const mins = parseInt(matchAny[2]);
      const period = matchAny[3].toUpperCase();

      if (period === 'PM' && hours !== 12) hours += 12;
      if (period === 'AM' && hours === 12) hours = 0;

      return hours * 60 + mins;
    }

    return 0;
  }

  // ============================================
  // HELPER: Minutes from midnight to time string
  // ============================================

  function minutesToTimeStr(mins) {
    const hours = Math.floor(mins / 60);
    const minutes = mins % 60;
    const period = hours >= 12 ? 'PM' : 'AM';
    const displayHours = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
    return `${displayHours}:${minutes.toString().padStart(2, '0')} ${period}`;
  }

  // ============================================
  // HELPER: Generate all 30-min slot times
  // ============================================

  function generateAllSlots() {
    const slots = [];
    const startMins = SLOT_CONFIG.START_HOUR * 60 + SLOT_CONFIG.START_MIN;
    const endMins = SLOT_CONFIG.END_HOUR * 60 + SLOT_CONFIG.END_MIN;

    for (let mins = startMins; mins <= endMins; mins += SLOT_CONFIG.SLOT_INTERVAL) {
      slots.push(minutesToTimeStr(mins));
    }
    return slots;
  }

  // ============================================
  // HELPER: Check if slot is during lunch
  // ============================================

  function isLunchTime(slotMins) {
    const lunchStart = SLOT_CONFIG.LUNCH_START * 60;
    const lunchEnd = SLOT_CONFIG.LUNCH_END * 60;
    return slotMins >= lunchStart && slotMins < lunchEnd;
  }

  // ============================================
  // HELPER: Check if service overlaps with lunch
  // ============================================

  function overlapsLunch(startMins, duration) {
    const endMins = startMins + duration;
    const lunchStart = SLOT_CONFIG.LUNCH_START * 60;
    const lunchEnd = SLOT_CONFIG.LUNCH_END * 60;
    return startMins < lunchEnd && endMins > lunchStart;
  }

  // ============================================
  // HELPER: Check if service exceeds cutoff
  // ============================================

  function exceedsCutoff(startMins, duration) {
    const endMins = startMins + duration + SLOT_CONFIG.TRAVEL_BUFFER;
    const cutoffMins = SLOT_CONFIG.HARD_CUTOFF * 60;
    return endMins > cutoffMins;
  }

  // ============================================
  // HELPER: Calculate distance (Haversine)
  // ============================================

  function calculateDistance(lat1, lng1, lat2, lng2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLng/2) * Math.sin(dLng/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
  }

  // ============================================
  // GET ADMIN DATA (Admin Panel)
  // ============================================

  function getAdminReports(phone) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);

    let adminsSheet = ss.getSheetByName(TABS.ADMINS);
    let isAdmin = false;
    let adminName = 'Admin';

    if (adminsSheet) {
      const adminsData = adminsSheet.getDataRange().getValues();
      for (let i = 1; i < adminsData.length; i++) {
        const adminPhone = String(adminsData[i][1] || '').replace(/\D/g, '').slice(-10);
        if (adminPhone === cleanPhone) {
          isAdmin = true;
          adminName = adminsData[i][0] || 'Admin';
          break;
        }
      }
    }

    if (!isAdmin) {
      return { success: false, error: 'Access denied. You are not registered as an admin.' };
    }

    const zonesSheet = ss.getSheetByName(TABS.GARDENER_ZONES);
    const zonesData = zonesSheet.getDataRange().getValues();
    const gardeners = [];
    const seenGardeners = {};

    for (let i = 1; i < zonesData.length; i++) {
      const gID = zonesData[i][2];
      const gName = zonesData[i][1];
      if (gID && gName && !seenGardeners[gID]) {
        gardeners.push({ id: gID, name: gName });
        seenGardeners[gID] = true;
      }
    }

    const today = new Date();
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();
    const bookings = [];

    for (let i = 1; i < bookingsData.length; i++) {
      const bookingDate = bookingsData[i][6];
      const bookingDateStr = parseDateToFormatted(bookingDate);

      if (bookingDateStr !== todayStr) continue;

      const isCompleted = checkIfJobCompleted(ss, bookingsData[i][0]);

      bookings.push({
        id: bookingsData[i][0],
        customerName: bookingsData[i][1],
        customerPhone: bookingsData[i][2],
        gardenerID: String(bookingsData[i][4] || ''),
        gardenerName: bookingsData[i][5],
        timeSlot: String(bookingsData[i][7] || ''),  // Convert to string (could be Date object)
        plantCount: bookingsData[i][8],
        address: bookingsData[i][9],
        mapLink: bookingsData[i][10],
        status: isCompleted ? 'COMPLETED' : 'PENDING'
      });
    }

    // Sort by time - convert to minutes for proper comparison
    bookings.sort((a, b) => {
      if (!a.timeSlot) return 1;
      if (!b.timeSlot) return -1;
      return timeToMinutes(a.timeSlot) - timeToMinutes(b.timeSlot);
    });

    return {
      success: true,
      adminName: adminName,
      bookings: bookings,
      gardeners: gardeners
    };
  }

  // ============================================
  // GET GARDENER JOBS (Partner App)
  // ============================================

  function getGardenerJobs(phone) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Clean input phone - extract last 10 digits
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);

    const zonesSheet = ss.getSheetByName(TABS.GARDENER_ZONES);
    const zonesData = zonesSheet.getDataRange().getValues();

    let gardenerID = null;
    let gardenerName = null;

    for (let i = 1; i < zonesData.length; i++) {
      // Clean sheet phone the same way - extract last 10 digits
      const gardenerPhone = String(zonesData[i][3] || '').replace(/\D/g, '').slice(-10);
      if (gardenerPhone === cleanPhone) {
        gardenerName = zonesData[i][1];
        gardenerID = zonesData[i][2];
        break;
      }
    }

    if (!gardenerID) {
      return { success: false, error: 'Gardener not found with this phone number' };
    }

    const today = new Date();
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();

    const jobs = [];

    for (let i = 1; i < bookingsData.length; i++) {
      const bookingGardenerID = String(bookingsData[i][4] || '');
      const bookingDate = bookingsData[i][6];

      if (bookingGardenerID !== String(gardenerID)) continue;

      const bookingDateStr = parseDateToFormatted(bookingDate);

      if (bookingDateStr !== todayStr) continue;

      const isCompleted = checkIfJobCompleted(ss, bookingsData[i][0]);

      jobs.push({
        id: bookingsData[i][0],
        customerName: bookingsData[i][1],
        customerPhone: bookingsData[i][2],
        address: bookingsData[i][9],
        mapLink: bookingsData[i][10],
        timeSlot: String(bookingsData[i][7] || ''),  // Convert to string (could be Date object)
        plantCount: bookingsData[i][8],
        notes: bookingsData[i][24] || '',  // Column Y - Partner Notes
        amount: bookingsData[i][20] || 0,  // Column U - Amount
        status: isCompleted ? 'COMPLETED' : 'PENDING'
      });
    }

    return {
      success: true,
      gardener: {
        id: gardenerID,
        name: gardenerName,
        phone: phone
      },
      jobs: jobs
    };
  }

  function checkIfJobCompleted(ss, bookingID) {
    const reportsSheet = ss.getSheetByName(TABS.SERVICE_REPORTS);
    if (!reportsSheet) return false;

    const reportsData = reportsSheet.getDataRange().getValues();
    for (let i = 1; i < reportsData.length; i++) {
      if (reportsData[i][1] === bookingID) {
        return true;
      }
    }
    return false;
  }

  // ============================================
  // SAVE SERVICE REPORT (Partner App)
  // ============================================

  function saveServiceReport(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let reportsSheet = ss.getSheetByName(TABS.SERVICE_REPORTS);
    if (!reportsSheet) {
      reportsSheet = ss.insertSheet(TABS.SERVICE_REPORTS);
      reportsSheet.appendRow([
        'ReportID',
        'BookingID',
        'GardenerID',
        'GardenerName',
        'GardenerPhone',
        'CustomerName',
        'CustomerPhone',
        'Address',
        'TotalPlants',
        'RedZonePlants',
        'Notes',
        'BeforePhotos',
        'AfterPhotos',
        'CompletedAt',
        'PaymentStatus',
        'PaymentLinkId',
        'Amount'
      ]);
    }

    const reportID = 'RPT' + Date.now().toString(36).toUpperCase();

    reportsSheet.appendRow([
      reportID,
      data.bookingID || '',
      data.gardenerID || '',
      data.gardenerName || '',
      data.gardenerPhone || '',
      data.customerName || '',
      data.customerPhone || '',
      data.address || '',
      data.totalPlants || '',
      data.redZonePlants || '',
      data.notes || '',
      data.beforePhotos || '',
      data.afterPhotos || '',
      new Date(),
      'Pending',
      '',
      data.amount || 0
    ]);

    // Send NPS form to customer after service completion
    if (data.customerPhone) {
      sendNPSForm(data.customerPhone);
      Logger.log('NPS form sent to: ' + data.customerPhone);
    }

    return {
      success: true,
      reportID: reportID,
      message: 'Service report saved successfully'
    };
  }

  // ============================================
  // GET AVAILABLE SLOTS (Booking Form)
  // ============================================

  function getAvailableSlots(pincode, customerLat, customerLng, plantCount) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const zonesSheet = ss.getSheetByName(TABS.GARDENER_ZONES);
    const zonesData = zonesSheet.getDataRange().getValues();

    // Calculate service duration based on plant count
    const serviceDuration = getServiceDuration(plantCount);

    let gardenerID = null;
    let gardenerName = null;

    if (customerLat && customerLng) {
      let closestDistance = Infinity;

      for (let i = 1; i < zonesData.length; i++) {
        const baseLat = parseFloat(zonesData[i][4]);
        const baseLng = parseFloat(zonesData[i][5]);

        if (!isNaN(baseLat) && !isNaN(baseLng)) {
          const distance = calculateDistance(customerLat, customerLng, baseLat, baseLng);

          if (distance < closestDistance) {
            closestDistance = distance;
            gardenerID = zonesData[i][2];
            gardenerName = zonesData[i][1];
          }
        }
      }
    }

    if (!gardenerID && pincode) {
      const inputPincode = String(pincode).trim().replace('.0', '');

      for (let i = 1; i < zonesData.length; i++) {
        const rawPincodes = String(zonesData[i][0] || '');
        const pincodes = rawPincodes.split(',').map(p => String(p).trim().replace('.0', ''));

        if (pincodes.includes(inputPincode)) {
          gardenerID = zonesData[i][2];
          gardenerName = zonesData[i][1];
          break;
        }
      }
    }

    if (!gardenerID) {
      return { available: false, message: 'Service not available in your area yet.' };
    }

    // Get all bookings for this gardener from Availability sheet
    const availSheet = ss.getSheetByName(TABS.AVAILABILITY);
    const availData = availSheet.getDataRange().getValues();

    // Also get bookings from Bookings sheet to know duration of each booking
    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();

    // Build a map of existing bookings with their durations per date
    // Format: { "dd/MM/yyyy": [{ startMins, duration }, ...] }
    const existingBookings = {};

    for (let i = 1; i < bookingsData.length; i++) {
      const bookingGardenerID = String(bookingsData[i][4] || '');
      if (bookingGardenerID !== String(gardenerID)) continue;

      const bookingDate = parseDateToFormatted(bookingsData[i][6]);
      const bookingTimeSlotRaw = bookingsData[i][7];
      const bookingPlantCount = bookingsData[i][8];

      if (!bookingDate || !bookingTimeSlotRaw) continue;

      // Handle time slot - could be Date object (TIME value) or string
      let bookingTimeSlot;
      if (bookingTimeSlotRaw instanceof Date ||
          (bookingTimeSlotRaw && typeof bookingTimeSlotRaw.getHours === 'function')) {
        // It's a TIME value stored as Date object - extract time in 12-hour format
        const hours = bookingTimeSlotRaw.getHours();
        const mins = bookingTimeSlotRaw.getMinutes();
        const period = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
        bookingTimeSlot = `${displayHours}:${mins.toString().padStart(2, '0')} ${period}`;
      } else {
        bookingTimeSlot = String(bookingTimeSlotRaw);
      }

      // Parse the time slot - handle both old format "8:30 AM - 10:00 AM" and new "8:30 AM"
      let startTimeStr = bookingTimeSlot;
      if (bookingTimeSlot.includes(' - ')) {
        startTimeStr = bookingTimeSlot.split(' - ')[0];
      }
      const startMins = timeToMinutes(startTimeStr);
      const duration = getServiceDuration(bookingPlantCount);

      if (!existingBookings[bookingDate]) {
        existingBookings[bookingDate] = [];
      }
      existingBookings[bookingDate].push({ startMins, duration });
    }

    // Generate all 30-min slot times
    const allSlots = generateAllSlots();

    // Helper: Check if a slot overlaps with any existing booking
    function overlapsExistingBooking(dateKey, slotStartMins, newDuration) {
      const dayBookings = existingBookings[dateKey] || [];
      const newEndMins = slotStartMins + newDuration + SLOT_CONFIG.TRAVEL_BUFFER;

      for (const booking of dayBookings) {
        const bookingEndMins = booking.startMins + booking.duration + SLOT_CONFIG.TRAVEL_BUFFER;

        // Check if there's any overlap
        if (slotStartMins < bookingEndMins && newEndMins > booking.startMins) {
          return true;
        }
      }
      return false;
    }

    // Generate dates with available slots
    const dates = [];
    for (let d = 1; d <= 60; d++) {
      const date = new Date();
      date.setDate(date.getDate() + d);
      const dateKey = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      const dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE, dd MMM');
      const dateInternal = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const times = [];

      for (const slotTime of allSlots) {
        const slotMins = timeToMinutes(slotTime);

        // Check 1: Is this during lunch? (1-2 PM)
        if (isLunchTime(slotMins)) {
          continue; // Skip lunch slots
        }

        // Check 2: Would service overlap with lunch?
        if (overlapsLunch(slotMins, serviceDuration)) {
          continue;
        }

        // Check 3: Would service exceed 6:30 PM cutoff?
        if (exceedsCutoff(slotMins, serviceDuration)) {
          continue;
        }

        // Check 4: Does this overlap with an existing booking?
        if (overlapsExistingBooking(dateKey, slotMins, serviceDuration)) {
          continue;
        }

        // Slot is available!
        times.push({
          time: slotTime,
          rowIndex: null
        });
      }

      dates.push({
        date: dateInternal,
        dateFormatted: dateFormatted,
        times: times
      });
    }

    return {
      available: true,
      gardenerName: gardenerName,
      gardenerID: gardenerID,
      serviceDuration: serviceDuration,
      dates: dates
    };
  }

  // ============================================
  // CREATE BOOKING (Booking Form)
  // ============================================

  function createBooking(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const required = ['pincode', 'gardenerID', 'date', 'timeSlot'];
    for (const field of required) {
      if (!data[field]) throw new Error('Missing: ' + field);
    }

    const [year, month, day] = data.date.split('-');
    const formattedDate = `${day}/${month}/${year}`;

    // === CRITICAL: Re-validate slot availability before booking ===
    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();

    const requestedSlotMins = timeToMinutes(data.timeSlot);
    const requestedDuration = getServiceDuration(data.plantCount);
    const requestedEndMins = requestedSlotMins + requestedDuration + SLOT_CONFIG.TRAVEL_BUFFER;

    for (let i = 1; i < bookingsData.length; i++) {
      const bookingGardenerID = String(bookingsData[i][4] || '');
      if (bookingGardenerID !== String(data.gardenerID)) continue;

      const bookingDate = parseDateToFormatted(bookingsData[i][6]);
      if (bookingDate !== formattedDate) continue;

      // This gardener already has a booking on this date - check for overlap
      // Handle time slot - could be Date object (TIME value) or string
      const existingTimeSlotRaw = bookingsData[i][7];
      let existingTimeSlot;
      if (existingTimeSlotRaw instanceof Date ||
          (existingTimeSlotRaw && typeof existingTimeSlotRaw.getHours === 'function')) {
        const hours = existingTimeSlotRaw.getHours();
        const mins = existingTimeSlotRaw.getMinutes();
        const period = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
        existingTimeSlot = `${displayHours}:${mins.toString().padStart(2, '0')} ${period}`;
      } else {
        existingTimeSlot = String(existingTimeSlotRaw || '');
      }

      let existingStartStr = existingTimeSlot;
      if (existingTimeSlot.includes(' - ')) {
        existingStartStr = existingTimeSlot.split(' - ')[0];
      }
      const existingStartMins = timeToMinutes(existingStartStr);
      const existingDuration = getServiceDuration(bookingsData[i][8]);
      const existingEndMins = existingStartMins + existingDuration + SLOT_CONFIG.TRAVEL_BUFFER;

      // Check overlap
      if (requestedSlotMins < existingEndMins && requestedEndMins > existingStartMins) {
        throw new Error('SLOT_TAKEN: This time slot is no longer available. Please select another slot.');
      }
    }
    // === END: Slot validation ===

    const bookingID = 'PP' + Date.now().toString(36).toUpperCase();

    bookingsSheet.appendRow([
      bookingID,
      data.customerName || 'Customer',
      data.phone || '',
      data.pincode,
      data.gardenerID,
      data.gardenerName || '',
      formattedDate,
      data.timeSlot,
      data.plantCount || '',
      data.address || data.fullAddress || '',
      data.mapLink || '',
      data.notes || '',
      new Date()
    ]);

    const availSheet = ss.getSheetByName(TABS.AVAILABILITY);

    if (data.availabilityRowIndex) {
      availSheet.getRange(data.availabilityRowIndex, 4).setValue(true);
    } else {
      availSheet.appendRow([
        data.gardenerID,
        formattedDate,
        data.timeSlot,
        true
      ]);
    }

    if (data.phone) {
      sendBookingConfirmation(
        data.phone,
        data.dateFormatted || formattedDate,
        data.timeSlot
      );
    }

    return {
      bookingID: bookingID,
      message: 'Booking confirmed! Visit on ' + formattedDate + ' at ' + data.timeSlot
    };
  }

  // ============================================
  // TEST FUNCTIONS
  // ============================================

  // RUN THIS TO VERIFY DUPLICATE DETECTION WORKS WITH YOUR ACTUAL DATA
  function testDuplicateDetection() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();

    Logger.log('=== TESTING DUPLICATE DETECTION ===');
    Logger.log('Total bookings: ' + (bookingsData.length - 1));

    // Group bookings by gardener + date
    const bookingsByGardenerDate = {};

    for (let i = 1; i < bookingsData.length; i++) {
      const gardenerID = String(bookingsData[i][4] || '');
      const dateRaw = bookingsData[i][6];
      const timeRaw = bookingsData[i][7];
      const customerName = bookingsData[i][1];

      // Test date parsing
      const dateFormatted = parseDateToFormatted(dateRaw);
      Logger.log(`Row ${i+1}: Date raw type = ${typeof dateRaw}, isDate = ${dateRaw instanceof Date}, formatted = ${dateFormatted}`);

      // Test time parsing
      let timeFormatted;
      if (timeRaw instanceof Date || (timeRaw && typeof timeRaw.getHours === 'function')) {
        const hours = timeRaw.getHours();
        const mins = timeRaw.getMinutes();
        const period = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
        timeFormatted = `${displayHours}:${mins.toString().padStart(2, '0')} ${period}`;
        Logger.log(`Row ${i+1}: Time was Date object, converted to: ${timeFormatted}`);
      } else {
        timeFormatted = String(timeRaw);
        Logger.log(`Row ${i+1}: Time was string: ${timeFormatted}`);
      }

      const startMins = timeToMinutes(timeFormatted.includes(' - ') ? timeFormatted.split(' - ')[0] : timeFormatted);
      Logger.log(`Row ${i+1}: startMins = ${startMins}`);

      const key = `${gardenerID}_${dateFormatted}`;
      if (!bookingsByGardenerDate[key]) {
        bookingsByGardenerDate[key] = [];
      }
      bookingsByGardenerDate[key].push({
        row: i + 1,
        customer: customerName,
        time: timeFormatted,
        startMins: startMins
      });
    }

    // Find duplicates
    Logger.log('\n=== CHECKING FOR OVERLAPS ===');
    let duplicatesFound = 0;

    for (const key in bookingsByGardenerDate) {
      const bookings = bookingsByGardenerDate[key];
      if (bookings.length > 1) {
        Logger.log(`\n${key} has ${bookings.length} bookings:`);

        // Check each pair for overlap
        for (let i = 0; i < bookings.length; i++) {
          for (let j = i + 1; j < bookings.length; j++) {
            const b1 = bookings[i];
            const b2 = bookings[j];
            const duration = 90; // default
            const buffer = 30;

            const b1End = b1.startMins + duration + buffer;
            const b2End = b2.startMins + duration + buffer;

            const overlaps = b1.startMins < b2End && b2.startMins < b1End;

            if (overlaps) {
              duplicatesFound++;
              Logger.log(`  ⚠️ OVERLAP: Row ${b1.row} (${b1.customer} @ ${b1.time}) vs Row ${b2.row} (${b2.customer} @ ${b2.time})`);
            }
          }
        }
      }
    }

    Logger.log(`\n=== RESULT: ${duplicatesFound} overlapping bookings found ===`);
    return duplicatesFound;
  }

  function testGetSlots() {
    // Test with 20 plants (1 hour service)
    Logger.log('=== Testing with 20 plants (1hr service) ===');
    const slots20 = getAvailableSlots('560102', 12.9716, 77.5946, 20);
    Logger.log('Service Duration: ' + slots20.serviceDuration + ' minutes');
    Logger.log('First date slots: ' + JSON.stringify(slots20.dates[0].times.map(t => t.time)));

    // Test with 40 plants (2 hour service)
    Logger.log('=== Testing with 40 plants (2hr service) ===');
    const slots40 = getAvailableSlots('560102', 12.9716, 77.5946, 40);
    Logger.log('Service Duration: ' + slots40.serviceDuration + ' minutes');
    Logger.log('First date slots: ' + JSON.stringify(slots40.dates[0].times.map(t => t.time)));

    // Test with 80 plants (3 hour service)
    Logger.log('=== Testing with 80 plants (3hr service) ===');
    const slots80 = getAvailableSlots('560102', 12.9716, 77.5946, 80);
    Logger.log('Service Duration: ' + slots80.serviceDuration + ' minutes');
    Logger.log('First date slots: ' + JSON.stringify(slots80.dates[0].times.map(t => t.time)));
  }

  function testSlotHelpers() {
    Logger.log('=== Testing Helper Functions ===');
    Logger.log('Duration for 15 plants: ' + getServiceDuration(15) + ' min');
    Logger.log('Duration for 30 plants: ' + getServiceDuration(30) + ' min');
    Logger.log('Duration for 45 plants: ' + getServiceDuration(45) + ' min');
    Logger.log('Duration for 60 plants: ' + getServiceDuration(60) + ' min');

    Logger.log('timeToMinutes("8:30 AM"): ' + timeToMinutes('8:30 AM'));
    Logger.log('timeToMinutes("1:00 PM"): ' + timeToMinutes('1:00 PM'));
    Logger.log('timeToMinutes("6:00 PM"): ' + timeToMinutes('6:00 PM'));

    Logger.log('minutesToTimeStr(510): ' + minutesToTimeStr(510)); // 8:30 AM
    Logger.log('minutesToTimeStr(780): ' + minutesToTimeStr(780)); // 1:00 PM
    Logger.log('minutesToTimeStr(1080): ' + minutesToTimeStr(1080)); // 6:00 PM

    Logger.log('All slots: ' + JSON.stringify(generateAllSlots()));
  }

  // ========================================
  // TRIGGER SETUP (Run once after updating)
  // ========================================

  function setupDailyFollowupTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'autoSendFiveDayFollowups' ||
          triggers[i].getHandlerFunction() === 'sendFiveDayFollowups') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    ScriptApp.newTrigger('autoSendFiveDayFollowups')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .inTimezone('Asia/Kolkata')
      .create();

    Logger.log('✅ Daily trigger set! autoSendFiveDayFollowups will run every day at 9 AM IST');
  }
