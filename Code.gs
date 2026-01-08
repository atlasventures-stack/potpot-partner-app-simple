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

    if (dateValue instanceof Date) {
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

      if (!pincode && !lat) {
        return jsonResponse({ error: 'Pincode or coordinates required' });
      }

      const slots = getAvailableSlots(pincode, lat, lng);
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
        gardenerID: bookingsData[i][4],
        gardenerName: bookingsData[i][5],
        timeSlot: bookingsData[i][7],
        plantCount: bookingsData[i][8],
        address: bookingsData[i][9],
        mapLink: bookingsData[i][10],
        status: isCompleted ? 'COMPLETED' : 'PENDING'
      });
    }

    bookings.sort((a, b) => {
      if (!a.timeSlot) return 1;
      if (!b.timeSlot) return -1;
      return a.timeSlot.localeCompare(b.timeSlot);
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

    const zonesSheet = ss.getSheetByName(TABS.GARDENER_ZONES);
    const zonesData = zonesSheet.getDataRange().getValues();

    let gardenerID = null;
    let gardenerName = null;

    for (let i = 1; i < zonesData.length; i++) {
      const gardenerPhone = String(zonesData[i][3] || '').trim();
      if (gardenerPhone === phone) {
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
      const bookingGardenerID = bookingsData[i][4];
      const bookingDate = bookingsData[i][6];

      if (bookingGardenerID !== gardenerID) continue;

      const bookingDateStr = parseDateToFormatted(bookingDate);

      if (bookingDateStr !== todayStr) continue;

      const isCompleted = checkIfJobCompleted(ss, bookingsData[i][0]);

      jobs.push({
        id: bookingsData[i][0],
        customerName: bookingsData[i][1],
        customerPhone: bookingsData[i][2],
        address: bookingsData[i][9],
        mapLink: bookingsData[i][10],
        timeSlot: bookingsData[i][7],
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
        'CompletedAt'
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
      new Date()
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

  function getAvailableSlots(pincode, customerLat, customerLng) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const zonesSheet = ss.getSheetByName(TABS.GARDENER_ZONES);
    const zonesData = zonesSheet.getDataRange().getValues();

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

    const availSheet = ss.getSheetByName(TABS.AVAILABILITY);
    const availData = availSheet.getDataRange().getValues();

    const bookedSlots = {};
    for (let i = 1; i < availData.length; i++) {
      const slotGardenerID = availData[i][0];
      const slotDate = availData[i][1];
      const timeSlot = availData[i][2];
      const isBooked = availData[i][3];

      if (slotGardenerID === gardenerID && isBooked) {
        const dateKey = parseDateToFormatted(slotDate);
        bookedSlots[dateKey + '|' + timeSlot] = true;
      }
    }

    const defaultTimeSlots = [
      '8:30 AM - 10:00 AM',
      '10:00 AM - 11:30 AM',
      '11:30 AM - 1:00 PM',
      '1:30 PM - 3:00 PM',
      '3:00 PM - 4:30 PM',
      '4:30 PM - 6:00 PM'
    ];

    const dates = [];
    for (let d = 1; d <= 60; d++) {
      const date = new Date();
      date.setDate(date.getDate() + d);
      const dateKey = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      const dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE, dd MMM');
      const dateInternal = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const times = [];
      defaultTimeSlots.forEach(timeSlot => {
        if (!bookedSlots[dateKey + '|' + timeSlot]) {
          times.push({
            time: timeSlot,
            rowIndex: null
          });
        }
      });

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

    const bookingID = 'PP' + Date.now().toString(36).toUpperCase();

    const [year, month, day] = data.date.split('-');
    const formattedDate = `${day}/${month}/${year}`;

    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
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

  function testGetSlots() {
    Logger.log(JSON.stringify(getAvailableSlots('560102', 12.9716, 77.5946), null, 2));
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
