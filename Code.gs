// ============================================
  // POTPOT BOOKING SYSTEM - Code.gs
  // ============================================

  const TABS = {
    GARDENER_ZONES: 'GardenerZones',
    AVAILABILITY: 'Availability',
    BOOKINGS: 'Bookings',
    SERVICE_REPORTS: 'ServiceReports',
    ADMINS: 'Admins',
    LOGS: 'Logs',
    HOLIDAYS: 'Holidays',
    // User Authentication & Dashboard
    USERS: 'Users',
    OTP_STORE: 'OTPStore',
    USER_PLANTS: 'UserPlants'
  };

  // ============================================
  // LOGGING SYSTEM - Writes to Logs sheet
  // ============================================

  function logToSheet(level, action, message, details) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logsSheet = ss.getSheetByName(TABS.LOGS);

      // Create Logs sheet if it doesn't exist
      if (!logsSheet) {
        logsSheet = ss.insertSheet(TABS.LOGS);
        logsSheet.appendRow(['Timestamp', 'Level', 'Action', 'Message', 'Details']);
        logsSheet.setFrozenRows(1);
        // Format header
        logsSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4a4a4a').setFontColor('#ffffff');
      }

      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const detailsStr = details ? JSON.stringify(details) : '';

      logsSheet.appendRow([timestamp, level, action, message, detailsStr]);

      // Also log to Apps Script logger
      Logger.log(`[${level}] ${action}: ${message} ${detailsStr}`);
    } catch (e) {
      // If logging fails, at least log to Apps Script
      Logger.log('LOGGING FAILED: ' + e.toString());
      Logger.log(`[${level}] ${action}: ${message}`);
    }
  }

  function logInfo(action, message, details) {
    logToSheet('INFO', action, message, details);
  }

  function logError(action, message, details) {
    logToSheet('ERROR', action, message, details);
  }

  function logWarning(action, message, details) {
    logToSheet('WARNING', action, message, details);
  }

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

          // Update PaymentCompletedAt (Column X = 24)
          if (data.paymentCompletedAt) {
            reportsSheet.getRange(i + 1, 24).setValue(data.paymentCompletedAt);
          }

          // Schedule Amplitude "Payment Success" event (async - doesn't block)
          if (data.status === 'Paid' || !data.status) {
            trackPaymentSuccess(reportsData[i]);
          }

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
  // AMPLITUDE SERVER-SIDE TRACKING (Async via Triggers)
  // ============================================

  const AMPLITUDE_API_KEY = '17abf777767237620e6398f575a1fff4';

  // Schedule Amplitude event to fire async (doesn't block main operation)
  function scheduleAmplitudeEvent(eventData) {
    try {
      // Create a one-time trigger to run after 5 seconds
      const trigger = ScriptApp.newTrigger('processDelayedAmplitudeEvent')
        .timeBased()
        .after(5 * 1000) // 5 seconds
        .create();

      // Store event data with trigger ID
      const props = PropertiesService.getScriptProperties();
      props.setProperty('amp_' + trigger.getUniqueId(), JSON.stringify(eventData));

      Logger.log('Amplitude event scheduled: ' + eventData.event_type);
    } catch (e) {
      Logger.log('Failed to schedule Amplitude event: ' + e.toString());
    }
  }

  // Triggered function - sends event to Amplitude
  function processDelayedAmplitudeEvent(e) {
    try {
      const triggerId = e.triggerUid;
      const props = PropertiesService.getScriptProperties();
      const dataKey = 'amp_' + triggerId;
      const dataStr = props.getProperty(dataKey);

      if (!dataStr) {
        Logger.log('No Amplitude data found for trigger: ' + triggerId);
        return;
      }

      const eventData = JSON.parse(dataStr);

      // Send to Amplitude
      const payload = {
        api_key: AMPLITUDE_API_KEY,
        events: [{
          user_id: eventData.user_id,
          event_type: eventData.event_type,
          time: eventData.time || Date.now(),
          event_properties: eventData.event_properties || {}
        }]
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api2.amplitude.com/2/httpapi', options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        Logger.log('✅ Amplitude event sent: ' + eventData.event_type + ' for user: ' + eventData.user_id);
      } else {
        Logger.log('Amplitude API error: ' + responseCode + ' - ' + response.getContentText());
      }

      // Clean up stored data
      props.deleteProperty(dataKey);

      // Delete the trigger
      const triggers = ScriptApp.getProjectTriggers();
      for (const t of triggers) {
        if (t.getUniqueId() === triggerId) {
          ScriptApp.deleteTrigger(t);
          break;
        }
      }
    } catch (error) {
      Logger.log('processDelayedAmplitudeEvent error: ' + error.toString());
    }
  }

  // Schedule "Service Completed" event
  function trackServiceCompleted(data) {
    const normalizedPhone = String(data.customerPhone || '').replace(/\D/g, '').slice(-10);
    if (!normalizedPhone) return;

    // Calculate service duration if we have start time
    let durationMinutes = null;
    if (data.serviceStartedAt) {
      const startTime = new Date(data.serviceStartedAt);
      const endTime = new Date();
      durationMinutes = Math.round((endTime - startTime) / (1000 * 60));
    }

    scheduleAmplitudeEvent({
      user_id: normalizedPhone,
      event_type: 'Service Completed',
      time: Date.now(),
      event_properties: {
        booking_id: data.bookingID || '',
        gardener_id: data.gardenerID || '',
        gardener_name: data.gardenerName || '',
        total_plants: data.totalPlants || 0,
        red_zone_plants: data.redZonePlants || 0,
        bugs_found: data.bugsFound || 'No',
        repotted_plants: data.repottedPlants || '',
        service_started_at: data.serviceStartedAt || '',
        service_completed_at: new Date().toISOString(),
        service_duration_minutes: durationMinutes,
        amount: data.amount || 0
      }
    });
  }

  // Schedule "Payment Success" event
  function trackPaymentSuccess(reportRow) {
    // reportRow is the full row from ServiceReports sheet
    const customerPhone = String(reportRow[6] || ''); // Column G
    const normalizedPhone = customerPhone.replace(/\D/g, '').slice(-10);
    if (!normalizedPhone) return;

    scheduleAmplitudeEvent({
      user_id: normalizedPhone,
      event_type: 'Payment Success',
      time: Date.now(),
      event_properties: {
        booking_id: reportRow[1] || '',      // Column B
        gardener_id: reportRow[2] || '',     // Column C
        gardener_name: reportRow[3] || '',   // Column D
        total_plants: reportRow[8] || 0,     // Column I
        amount: reportRow[16] || 0,          // Column Q
        currency: 'INR',
        payment_method: 'Razorpay'
      }
    });
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

      // ============================================
      // USER AUTHENTICATION & DASHBOARD ENDPOINTS
      // ============================================

      if (action === 'sendOTP') {
        const phone = e.parameter.phone;
        if (!phone) {
          return jsonResponse({ success: false, error: 'Phone is required' });
        }
        const result = sendOTPToUser(phone);
        return jsonResponse(result);
      }

      if (action === 'verifyOTP') {
        const phone = e.parameter.phone;
        const otp = e.parameter.otp;
        if (!phone || !otp) {
          return jsonResponse({ success: false, error: 'Phone and OTP are required' });
        }
        const result = verifyUserOTP(phone, otp);
        return jsonResponse(result);
      }

      if (action === 'getCustomerBookings') {
        const phone = e.parameter.phone;
        if (!phone) {
          return jsonResponse({ success: false, error: 'Phone is required' });
        }
        const result = getCustomerBookings(phone);
        return jsonResponse(result);
      }

      if (action === 'getCustomerPlants') {
        const phone = e.parameter.phone;
        if (!phone) {
          return jsonResponse({ success: false, error: 'Phone is required' });
        }
        const result = getCustomerPlants(phone);
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

      if (data.action === 'saveCustomerPlant') {
        const result = saveCustomerPlant(data);
        return jsonResponse(result);
      }

      // This is a booking request
      const result = createBooking(data);
      return jsonResponse({ success: true, booking: result });
    } catch (error) {
      // LOG: API error
      logError('API_POST_ERROR', 'doPost threw error', {
        error: error.toString(),
        stack: error.stack || 'no stack',
        postData: e.postData ? e.postData.contents.substring(0, 500) : 'no postData'
      });
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
  // HELPER: Check if gardener is on holiday for a date
  // ============================================

  function isGardenerOnHoliday(ss, gardenerID, checkDate) {
    const holidaysSheet = ss.getSheetByName(TABS.HOLIDAYS);
    if (!holidaysSheet) return false;

    const holidaysData = holidaysSheet.getDataRange().getValues();
    const checkDateOnly = new Date(checkDate);
    checkDateOnly.setHours(0, 0, 0, 0);

    for (let i = 1; i < holidaysData.length; i++) {
      const holidayGardenerID = String(holidaysData[i][0] || '');
      if (holidayGardenerID !== String(gardenerID)) continue;

      const startDate = parseDate(holidaysData[i][1]);
      const endDate = parseDate(holidaysData[i][2]);

      if (!startDate) continue;

      // Set times to start of day for comparison
      startDate.setHours(0, 0, 0, 0);

      // If no end date, treat as single day holiday
      const effectiveEndDate = endDate ? new Date(endDate) : new Date(startDate);
      effectiveEndDate.setHours(23, 59, 59, 999);

      // Check if checkDate falls within holiday range
      if (checkDateOnly >= startDate && checkDateOnly <= effectiveEndDate) {
        return true;
      }
    }

    return false;
  }

  // ============================================
  // HELPER: Get holidays count for gardener in date range
  // ============================================

  function countHolidayDays(ss, gardenerID, startDate, endDate) {
    const holidaysSheet = ss.getSheetByName(TABS.HOLIDAYS);
    if (!holidaysSheet) return 0;

    const holidaysData = holidaysSheet.getDataRange().getValues();
    let holidayCount = 0;

    // Iterate through each day in the range
    const currentDate = new Date(startDate);
    currentDate.setHours(0, 0, 0, 0);
    const end = new Date(endDate);
    end.setHours(0, 0, 0, 0);

    while (currentDate <= end) {
      for (let i = 1; i < holidaysData.length; i++) {
        const holidayGardenerID = String(holidaysData[i][0] || '');
        if (holidayGardenerID !== String(gardenerID)) continue;

        const holStart = parseDate(holidaysData[i][1]);
        const holEnd = parseDate(holidaysData[i][2]);

        if (!holStart) continue;

        holStart.setHours(0, 0, 0, 0);
        const effectiveEnd = holEnd ? new Date(holEnd) : new Date(holStart);
        effectiveEnd.setHours(23, 59, 59, 999);

        if (currentDate >= holStart && currentDate <= effectiveEnd) {
          holidayCount++;
          break; // Don't double count same day
        }
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    return holidayCount;
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
  // HELPER: Format time slot for display
  // ============================================

  function formatTimeSlotForDisplay(timeSlotRaw) {
    if (!timeSlotRaw) return '';

    // Handle Date objects from Google Sheets
    if (timeSlotRaw instanceof Date ||
        (timeSlotRaw && typeof timeSlotRaw.getHours === 'function')) {
      const hours = timeSlotRaw.getHours();
      const mins = timeSlotRaw.getMinutes();
      const period = hours >= 12 ? 'PM' : 'AM';
      const displayHours = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
      return `${displayHours}:${mins.toString().padStart(2, '0')} ${period}`;
    }

    // Already a string - return as-is
    return String(timeSlotRaw);
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
        timeSlot: formatTimeSlotForDisplay(bookingsData[i][7]),
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
        timeSlot: formatTimeSlotForDisplay(bookingsData[i][7]),
        plantCount: bookingsData[i][8],
        notes: bookingsData[i][24] || '',  // Column Y - Partner Notes
        amount: bookingsData[i][20] !== '' && bookingsData[i][20] !== null && bookingsData[i][20] !== undefined ? bookingsData[i][20] : null,  // Column U - Amount (null if blank, so frontend calculates from plants)
        status: isCompleted ? 'COMPLETED' : 'PENDING'
      });
    }

    // Sort jobs by time - earliest first
    jobs.sort((a, b) => {
      if (!a.timeSlot) return 1;
      if (!b.timeSlot) return -1;
      return timeToMinutes(a.timeSlot) - timeToMinutes(b.timeSlot);
    });

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
  // NPS FORM DELAYED SENDER (One-time trigger)
  // ============================================

  function scheduleNPSForm(customerPhone, reportID) {
    try {
      // Create a one-time trigger to run after 2 minutes
      const trigger = ScriptApp.newTrigger('processDelayedNPS')
        .timeBased()
        .after(2 * 60 * 1000) // 2 minutes in milliseconds
        .create();

      // Store the data for this trigger using its unique ID
      const props = PropertiesService.getScriptProperties();
      props.setProperty('nps_' + trigger.getUniqueId(), JSON.stringify({
        phone: customerPhone,
        reportID: reportID
      }));

      Logger.log('NPS scheduled for: ' + customerPhone + ' (trigger in 2 min)');
    } catch (e) {
      Logger.log('Failed to schedule NPS: ' + e.toString());
      // Fallback: try to send immediately if trigger creation fails
      try {
        sendNPSForm(customerPhone);
      } catch (e2) {
        Logger.log('Fallback NPS also failed: ' + e2.toString());
      }
    }
  }

  function processDelayedNPS(e) {
    try {
      // Get trigger ID from the event
      const triggerId = e.triggerUid;
      const props = PropertiesService.getScriptProperties();
      const dataKey = 'nps_' + triggerId;
      const dataStr = props.getProperty(dataKey);

      if (!dataStr) {
        Logger.log('No NPS data found for trigger: ' + triggerId);
        return;
      }

      const data = JSON.parse(dataStr);

      // Send the NPS form
      sendNPSForm(data.phone);
      Logger.log('✅ NPS sent to: ' + data.phone + ' (Report: ' + data.reportID + ')');

      // Clean up: remove the stored data
      props.deleteProperty(dataKey);

      // Clean up: delete this trigger (it's one-time but let's be safe)
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getUniqueId() === triggerId) {
          ScriptApp.deleteTrigger(trigger);
          break;
        }
      }
    } catch (error) {
      Logger.log('❌ processDelayedNPS error: ' + error.toString());
    }
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
        'Amount',
        'BugsFound',
        'RepottedPlants',
        'ServiceStartedAt',
        'BeforePhotosAt',
        'AfterPhotosAt',
        'PaymentCompletedAt'
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
      data.amount || 0,
      data.bugsFound || 'No',
      data.repottedPlants || '',
      data.serviceStartedAt || '',
      data.beforePhotosAt || '',
      data.afterPhotosAt || '',
      ''  // PaymentCompletedAt - filled when payment is done
    ]);

    // Schedule Amplitude "Service Completed" event (async - doesn't block)
    if (data.customerPhone) {
      trackServiceCompleted(data);
    }

    // Schedule NPS form to send in background (gardener doesn't wait)
    if (data.customerPhone) {
      scheduleNPSForm(data.customerPhone, reportID);
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

      // Find ALL gardeners matching this pincode
      const matchingGardeners = [];
      for (let i = 1; i < zonesData.length; i++) {
        const rawPincodes = String(zonesData[i][0] || '');
        const pincodes = rawPincodes.split(',').map(p => String(p).trim().replace('.0', ''));

        if (pincodes.includes(inputPincode)) {
          matchingGardeners.push({
            id: zonesData[i][2],
            name: zonesData[i][1]
          });
        }
      }

      if (matchingGardeners.length === 1) {
        // Only one gardener - use them
        gardenerID = matchingGardeners[0].id;
        gardenerName = matchingGardeners[0].name;
      } else if (matchingGardeners.length > 1) {
        // Multiple gardeners - load balance by counting bookings + holidays
        const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
        const bookingsData = bookingsSheet.getDataRange().getValues();

        // Count bookings + holidays for each gardener in the next 7 days
        const today = new Date();
        const next7Days = new Date();
        next7Days.setDate(today.getDate() + 7);

        let minBlockedDays = Infinity;
        let selectedGardener = matchingGardeners[0]; // Default to first

        for (const gardener of matchingGardeners) {
          let bookingCount = 0;

          for (let i = 1; i < bookingsData.length; i++) {
            const bookingGardenerID = String(bookingsData[i][4] || '');
            if (bookingGardenerID !== String(gardener.id)) continue;

            const bookingDate = parseDate(bookingsData[i][6]);
            if (!bookingDate) continue;

            // Only count future bookings (next 7 days)
            if (bookingDate >= today && bookingDate <= next7Days) {
              bookingCount++;
            }
          }

          // Also count holiday days in next 7 days
          const holidayCount = countHolidayDays(ss, gardener.id, today, next7Days);
          const totalBlocked = bookingCount + holidayCount;

          logInfo('LOAD_BALANCE', `Gardener ${gardener.id} (${gardener.name}): ${bookingCount} bookings + ${holidayCount} holidays = ${totalBlocked} blocked days`);

          if (totalBlocked < minBlockedDays) {
            minBlockedDays = totalBlocked;
            selectedGardener = gardener;
          }
        }

        gardenerID = selectedGardener.id;
        gardenerName = selectedGardener.name;
        logInfo('LOAD_BALANCE', `Selected gardener: ${gardenerID} (${gardenerName}) with ${minBlockedDays} blocked days`);
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

      // Check if gardener is on holiday for this date
      if (isGardenerOnHoliday(ss, gardenerID, date)) {
        // Gardener is on holiday - no slots available for this date
        dates.push({
          date: dateInternal,
          dateFormatted: dateFormatted,
          times: [],
          holiday: true
        });
        continue;
      }

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

    // LOG: Booking attempt started
    logInfo('BOOKING_START', 'New booking attempt', {
      customer: data.customerName,
      phone: data.phone,
      date: data.date,
      timeSlot: data.timeSlot,
      pincode: data.pincode,
      gardenerID: data.gardenerID,
      plantCount: data.plantCount
    });

    const required = ['pincode', 'gardenerID', 'date', 'timeSlot'];
    for (const field of required) {
      if (!data[field]) {
        logError('BOOKING_VALIDATION', 'Missing required field: ' + field, data);
        throw new Error('Missing: ' + field);
      }
    }

    const [year, month, day] = data.date.split('-');
    const formattedDate = `${day}/${month}/${year}`;

    // === CRITICAL: Re-validate slot availability before booking ===
    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();

    logInfo('BOOKING_SLOT_CHECK', 'Checking slot availability', {
      date: formattedDate,
      timeSlot: data.timeSlot,
      gardenerID: data.gardenerID,
      existingBookingsCount: bookingsData.length - 1
    });

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
        logWarning('BOOKING_SLOT_TAKEN', 'Slot already taken', {
          requestedSlot: data.timeSlot,
          existingBookingID: bookingsData[i][0],
          existingSlot: existingTimeSlot,
          customer: data.customerName
        });
        throw new Error('SLOT_TAKEN: This time slot is no longer available. Please select another slot.');
      }
    }
    // === END: Slot validation ===

    logInfo('BOOKING_SLOT_OK', 'Slot available, proceeding with save', {
      date: formattedDate,
      timeSlot: data.timeSlot
    });

    const bookingID = 'PP' + Date.now().toString(36).toUpperCase();

    // Prepare booking row data
    // Columns A-M written by code, N-AD filled by other processes, AE=LeadSource (new)
    const bookingRow = [
      bookingID,                              // 0  - A: BookingID
      data.customerName || 'Customer',        // 1  - B: CustomerName
      data.phone || '',                       // 2  - C: Phone
      data.pincode,                           // 3  - D: Pincode
      data.gardenerID,                        // 4  - E: GardenerID
      data.gardenerName || '',                // 5  - F: GardenerName
      formattedDate,                          // 6  - G: Date
      data.timeSlot,                          // 7  - H: TimeSlot
      data.plantCount || '',                  // 8  - I: PlantCount
      data.address || data.fullAddress || '', // 9  - J: Address
      data.mapLink || '',                     // 10 - K: MapLink
      '',                                     // 11 - L: ReachTime
      new Date(),                             // 12 - M: BookedAt
      '', '', '', '', '', '', '', '',         // 13-20: N-U (Notes through Amount)
      '', '', '', '', '', '', '', '', '',     // 21-29: V-AD (Payment through col AD)
      data.leadSource || 'Unknown'            // 30 - AE: LeadSource
    ];

    // LOG: About to save
    logInfo('BOOKING_SAVE_START', 'Attempting to save booking', {
      bookingID: bookingID,
      rowLength: bookingRow.length
    });

    // Save to sheet - using getLastRow + setValues (more reliable than appendRow with protected sheets)
    let targetRow;
    try {
      const lastRow = bookingsSheet.getLastRow();
      targetRow = lastRow + 1;

      logInfo('BOOKING_WRITE_START', 'Writing to specific row', {
        bookingID: bookingID,
        lastRow: lastRow,
        targetRow: targetRow,
        columns: bookingRow.length
      });

      // Use setValues instead of appendRow - more reliable with protected sheets
      bookingsSheet.getRange(targetRow, 1, 1, bookingRow.length).setValues([bookingRow]);

      logInfo('BOOKING_WRITE_OK', 'setValues completed', { bookingID: bookingID, row: targetRow });
    } catch (writeError) {
      logError('BOOKING_WRITE_FAIL', 'setValues threw error', {
        bookingID: bookingID,
        targetRow: targetRow,
        error: writeError.toString(),
        errorStack: writeError.stack || 'no stack'
      });
      throw writeError;
    }

    // CRITICAL: Force write to complete
    try {
      SpreadsheetApp.flush();
      logInfo('BOOKING_FLUSH_OK', 'flush() completed', { bookingID: bookingID });
    } catch (flushError) {
      logError('BOOKING_FLUSH_FAIL', 'flush() threw error', {
        bookingID: bookingID,
        error: flushError.toString()
      });
    }

    // Small delay to ensure persistence
    Utilities.sleep(500); // Increased from 300ms to 500ms

    // VERIFY the row was actually written
    logInfo('BOOKING_VERIFY_START', 'Starting verification', { bookingID: bookingID });

    const verifyData = bookingsSheet.getDataRange().getValues();
    const totalRows = verifyData.length;
    let bookingFound = false;
    let foundAtRow = -1;

    for (let i = verifyData.length - 1; i >= Math.max(1, verifyData.length - 10); i--) {
      if (verifyData[i][0] === bookingID) {
        bookingFound = true;
        foundAtRow = i + 1;
        break;
      }
    }

    logInfo('BOOKING_VERIFY_RESULT', 'First verification result', {
      bookingID: bookingID,
      found: bookingFound,
      foundAtRow: foundAtRow,
      totalRows: totalRows,
      lastRowsChecked: Math.min(10, totalRows - 1)
    });

    // If not found, retry once with setValues
    if (!bookingFound) {
      logWarning('BOOKING_RETRY', 'Booking not found, attempting retry with setValues', { bookingID: bookingID });

      try {
        const retryLastRow = bookingsSheet.getLastRow();
        const retryTargetRow = retryLastRow + 1;
        logInfo('BOOKING_RETRY_WRITE', 'Retry writing to row', { bookingID: bookingID, targetRow: retryTargetRow });

        bookingsSheet.getRange(retryTargetRow, 1, 1, bookingRow.length).setValues([bookingRow]);
        logInfo('BOOKING_RETRY_WRITE_OK', 'Retry setValues completed', { bookingID: bookingID, row: retryTargetRow });
      } catch (retryWriteError) {
        logError('BOOKING_RETRY_WRITE_FAIL', 'Retry setValues failed', {
          bookingID: bookingID,
          error: retryWriteError.toString(),
          errorStack: retryWriteError.stack || 'no stack'
        });
      }

      SpreadsheetApp.flush();
      Utilities.sleep(500);

      // Check again
      const retryData = bookingsSheet.getDataRange().getValues();
      for (let i = retryData.length - 1; i >= Math.max(1, retryData.length - 10); i--) {
        if (retryData[i][0] === bookingID) {
          bookingFound = true;
          foundAtRow = i + 1;
          break;
        }
      }

      logInfo('BOOKING_RETRY_RESULT', 'Retry verification result', {
        bookingID: bookingID,
        found: bookingFound,
        foundAtRow: foundAtRow,
        totalRowsAfterRetry: retryData.length
      });
    }

    // FINAL LOG: Success or failure
    if (bookingFound) {
      logInfo('BOOKING_SUCCESS', 'Booking saved and verified', {
        bookingID: bookingID,
        customer: data.customerName,
        phone: data.phone,
        date: formattedDate,
        timeSlot: data.timeSlot,
        row: foundAtRow
      });
    } else {
      logError('BOOKING_FAILED', 'CRITICAL: Booking could not be verified after retry', {
        bookingID: bookingID,
        customer: data.customerName,
        phone: data.phone,
        date: formattedDate,
        timeSlot: data.timeSlot,
        address: data.address || data.fullAddress
      });
    }

    // Update availability (fast, same spreadsheet)
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

    // Schedule async notifications (email + WhatsApp) - don't block response
    scheduleBookingNotifications({
      bookingID: bookingID,
      bookingFound: bookingFound,
      foundAtRow: foundAtRow,
      customerName: data.customerName,
      phone: data.phone,
      formattedDate: formattedDate,
      timeSlot: data.timeSlot,
      plantCount: data.plantCount,
      address: data.address || data.fullAddress,
      gardenerName: data.gardenerName,
      gardenerID: data.gardenerID,
      mapLink: data.mapLink
    });

    // Return immediately - notifications will be sent async
    return {
      bookingID: bookingID,
      message: 'Booking confirmed! Visit on ' + formattedDate + ' at ' + data.timeSlot
    };
  }

  // ============================================
  // ASYNC BOOKING NOTIFICATIONS (Email + WhatsApp)
  // ============================================

  function scheduleBookingNotifications(notificationData) {
    try {
      // Store notification data in Properties for the trigger to access
      const props = PropertiesService.getScriptProperties();
      const key = 'BOOKING_NOTIFY_' + notificationData.bookingID;
      props.setProperty(key, JSON.stringify(notificationData));

      // Create a one-time trigger to send notifications in 2 seconds
      ScriptApp.newTrigger('processBookingNotifications')
        .timeBased()
        .after(2000) // 2 seconds
        .create();

      logInfo('BOOKING_NOTIFY_SCHEDULED', 'Notifications scheduled', { bookingID: notificationData.bookingID });
    } catch (e) {
      // If scheduling fails, send synchronously as fallback
      logError('BOOKING_NOTIFY_SCHEDULE_FAIL', 'Failed to schedule, sending sync', { error: e.toString() });
      sendBookingNotificationsSync(notificationData);
    }
  }

  function processBookingNotifications(e) {
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();

    // Find and process booking notification
    for (const key in allProps) {
      if (key.startsWith('BOOKING_NOTIFY_')) {
        try {
          const data = JSON.parse(allProps[key]);
          sendBookingNotificationsSync(data);
          props.deleteProperty(key);
        } catch (err) {
          logError('BOOKING_NOTIFY_PROCESS_FAIL', 'Failed to process notification', { key: key, error: err.toString() });
          props.deleteProperty(key);
        }
      }
    }

    // Clean up the trigger
    if (e && e.triggerUid) {
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getUniqueId() === e.triggerUid) {
          ScriptApp.deleteTrigger(trigger);
          break;
        }
      }
    }
  }

  function sendBookingNotificationsSync(data) {
    // Send EMAIL backup
    try {
      MailApp.sendEmail({
        to: 'potpot@atlasventuresonline.com',
        subject: data.bookingFound ? '✅ PotPot Booking: ' + data.bookingID : '❌ BOOKING FAILED: ' + data.bookingID,
        body: 'Booking ID: ' + data.bookingID +
              '\nCustomer: ' + (data.customerName || 'N/A') +
              '\nPhone: ' + (data.phone || 'N/A') +
              '\nDate: ' + data.formattedDate +
              '\nTime: ' + data.timeSlot +
              '\nPlants: ' + (data.plantCount || 'N/A') +
              '\nAddress: ' + (data.address || 'N/A') +
              '\nGardener: ' + (data.gardenerName || 'N/A') + ' (' + data.gardenerID + ')' +
              '\nMap: ' + (data.mapLink || 'N/A') +
              '\n\n' + (data.bookingFound ? '✅ Verified in sheet: YES (Row ' + data.foundAtRow + ')' : '❌ VERIFIED IN SHEET: NO - CHECK LOGS TAB!')
      });
      logInfo('BOOKING_EMAIL_SENT', 'Backup email sent', { bookingID: data.bookingID });
    } catch (emailError) {
      logError('BOOKING_EMAIL_FAIL', 'Failed to send backup email', {
        bookingID: data.bookingID,
        error: emailError.toString()
      });
    }

    // Send WHATSAPP confirmation
    if (data.phone) {
      try {
        sendBookingConfirmation(data.phone, data.formattedDate, data.timeSlot);
        logInfo('BOOKING_WHATSAPP_SENT', 'WhatsApp confirmation sent', {
          bookingID: data.bookingID,
          phone: data.phone
        });
      } catch (whatsappError) {
        logError('BOOKING_WHATSAPP_FAIL', 'WhatsApp confirmation failed', {
          bookingID: data.bookingID,
          phone: data.phone,
          error: whatsappError.toString()
        });
      }
    }
  }

  // ============================================
  // USER AUTHENTICATION FUNCTIONS
  // ============================================

  // Send OTP to user via WATI WhatsApp
  function sendOTPToUser(phone) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);

    if (cleanPhone.length !== 10) {
      return { success: false, error: 'Invalid phone number' };
    }

    // Generate 4-digit OTP
    const otp = String(Math.floor(1000 + Math.random() * 9000));
    const now = new Date();
    const expiresAt = new Date(now.getTime() + 5 * 60 * 1000); // 5 minutes

    // Get or create OTPStore sheet
    let otpSheet = ss.getSheetByName(TABS.OTP_STORE);
    if (!otpSheet) {
      otpSheet = ss.insertSheet(TABS.OTP_STORE);
      otpSheet.appendRow(['Phone', 'OTP', 'CreatedAt', 'ExpiresAt', 'Used']);
      otpSheet.setFrozenRows(1);
    }

    // Store OTP
    otpSheet.appendRow([cleanPhone, otp, now, expiresAt, false]);

    // Send OTP via WATI WhatsApp
    try {
      const url = `${WATI_API_ENDPOINT}/api/v1/sendTemplateMessage?whatsappNumber=91${cleanPhone}`;

      const payload = {
        "template_name": "otp_verification",
        "broadcast_name": "otp_" + Date.now(),
        "parameters": [
          { "name": "1", "value": otp }
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

      const response = UrlFetchApp.fetch(url, options);
      const result = JSON.parse(response.getContentText());

      if (result.result === true) {
        logInfo('OTP_SENT', 'OTP sent successfully', { phone: cleanPhone });
        return { success: true, message: 'OTP sent to your WhatsApp' };
      } else {
        logError('OTP_SEND_FAIL', 'WATI API returned error', { phone: cleanPhone, result: result });
        return { success: false, error: 'Failed to send OTP. Please try again.' };
      }
    } catch (error) {
      logError('OTP_SEND_ERROR', 'Exception sending OTP', { phone: cleanPhone, error: error.toString() });
      return { success: false, error: 'Failed to send OTP. Please try again.' };
    }
  }

  // Verify OTP and create/update user
  function verifyUserOTP(phone, enteredOTP) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);

    const otpSheet = ss.getSheetByName(TABS.OTP_STORE);
    if (!otpSheet) {
      return { success: false, error: 'No OTP found. Please request a new one.' };
    }

    const otpData = otpSheet.getDataRange().getValues();
    const now = new Date();

    // Find the latest unused OTP for this phone
    let validOTP = null;
    let otpRowIndex = -1;

    for (let i = otpData.length - 1; i >= 1; i--) {
      const otpPhone = String(otpData[i][0]);
      const otp = String(otpData[i][1]);
      const expiresAt = otpData[i][3];
      const used = otpData[i][4];

      if (otpPhone === cleanPhone && !used) {
        // Check if expired
        const expiryDate = expiresAt instanceof Date ? expiresAt : new Date(expiresAt);
        if (now > expiryDate) {
          continue; // Expired, skip
        }

        validOTP = otp;
        otpRowIndex = i + 1;
        break;
      }
    }

    if (!validOTP) {
      return { success: false, error: 'OTP expired or not found. Please request a new one.' };
    }

    if (validOTP !== String(enteredOTP)) {
      return { success: false, error: 'Invalid OTP. Please try again.' };
    }

    // Mark OTP as used
    otpSheet.getRange(otpRowIndex, 5).setValue(true);

    // Create or update user in Users sheet
    let usersSheet = ss.getSheetByName(TABS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(TABS.USERS);
      usersSheet.appendRow(['Phone', 'Name', 'CreatedAt', 'LastLoginAt']);
      usersSheet.setFrozenRows(1);
    }

    const usersData = usersSheet.getDataRange().getValues();
    let userRowIndex = -1;
    let userName = '';
    let isNewUser = true;

    for (let i = 1; i < usersData.length; i++) {
      if (String(usersData[i][0]) === cleanPhone) {
        userRowIndex = i + 1;
        userName = usersData[i][1] || '';
        isNewUser = false;
        break;
      }
    }

    if (isNewUser) {
      // Try to get name from Bookings sheet
      const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
      if (bookingsSheet) {
        const bookingsData = bookingsSheet.getDataRange().getValues();
        for (let i = bookingsData.length - 1; i >= 1; i--) {
          const bookingPhone = String(bookingsData[i][2]).replace(/\D/g, '').slice(-10);
          if (bookingPhone === cleanPhone) {
            userName = bookingsData[i][1] || '';
            break;
          }
        }
      }

      // Create new user
      usersSheet.appendRow([cleanPhone, userName, now, now]);
      logInfo('USER_CREATED', 'New user registered', { phone: cleanPhone, name: userName });
    } else {
      // Update last login
      usersSheet.getRange(userRowIndex, 4).setValue(now);
      logInfo('USER_LOGIN', 'User logged in', { phone: cleanPhone, name: userName });
    }

    return {
      success: true,
      user: {
        phone: cleanPhone,
        name: userName,
        isNewUser: isNewUser
      }
    };
  }

  // ============================================
  // USER DASHBOARD FUNCTIONS
  // ============================================

  // Get all bookings for a customer
  function getCustomerBookings(phone) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);

    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    if (!bookingsSheet) {
      return { success: true, bookings: [], upcomingCount: 0, pastCount: 0 };
    }

    const bookingsData = bookingsSheet.getDataRange().getValues();
    const reportsSheet = ss.getSheetByName(TABS.SERVICE_REPORTS);
    const reportsData = reportsSheet ? reportsSheet.getDataRange().getValues() : [];

    // Build set of completed booking IDs
    const completedBookings = new Set();
    for (let i = 1; i < reportsData.length; i++) {
      completedBookings.add(reportsData[i][1]); // BookingID is at index 1
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const bookings = [];

    for (let i = 1; i < bookingsData.length; i++) {
      const bookingPhone = String(bookingsData[i][2]).replace(/\D/g, '').slice(-10);

      if (bookingPhone !== cleanPhone) continue;

      const bookingID = bookingsData[i][0];
      const bookingDate = parseDate(bookingsData[i][6]);

      if (!bookingDate) continue;

      bookingDate.setHours(0, 0, 0, 0);

      let status;
      if (completedBookings.has(bookingID)) {
        status = 'COMPLETED';
      } else if (bookingDate >= today) {
        status = 'UPCOMING';
      } else {
        status = 'PAST';
      }

      bookings.push({
        id: bookingID,
        customerName: bookingsData[i][1],
        date: parseDateToFormatted(bookingsData[i][6]),
        dateFormatted: Utilities.formatDate(bookingDate, Session.getScriptTimeZone(), 'EEE, dd MMM yyyy'),
        timeSlot: formatTimeSlotForDisplay(bookingsData[i][7]),
        plantCount: bookingsData[i][8],
        address: bookingsData[i][9],
        gardenerName: bookingsData[i][5],
        status: status
      });
    }

    // Sort: Upcoming first (by date asc), then Past (by date desc)
    bookings.sort((a, b) => {
      if (a.status === 'UPCOMING' && b.status !== 'UPCOMING') return -1;
      if (a.status !== 'UPCOMING' && b.status === 'UPCOMING') return 1;

      const dateA = parseDate(a.date);
      const dateB = parseDate(b.date);

      if (a.status === 'UPCOMING') {
        return dateA - dateB; // Ascending for upcoming
      } else {
        return dateB - dateA; // Descending for past/completed
      }
    });

    const upcomingCount = bookings.filter(b => b.status === 'UPCOMING').length;
    const pastCount = bookings.filter(b => b.status !== 'UPCOMING').length;

    return {
      success: true,
      bookings: bookings,
      upcomingCount: upcomingCount,
      pastCount: pastCount
    };
  }

  // Get saved plant analyses for a customer
  function getCustomerPlants(phone) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);

    const plantsSheet = ss.getSheetByName(TABS.USER_PLANTS);
    if (!plantsSheet) {
      return { success: true, plants: [], schedules: [], recommendations: [] };
    }

    const plantsData = plantsSheet.getDataRange().getValues();

    const schedules = [];
    const recommendations = [];

    for (let i = 1; i < plantsData.length; i++) {
      const plantPhone = String(plantsData[i][1]).replace(/\D/g, '').slice(-10);

      if (plantPhone !== cleanPhone) continue;

      const entry = {
        id: plantsData[i][0],
        type: plantsData[i][2],
        imageURLs: plantsData[i][3] ? String(plantsData[i][3]).split(',') : [],
        data: plantsData[i][4] ? JSON.parse(plantsData[i][4]) : {},
        plantCount: plantsData[i][5] || 0,
        createdAt: plantsData[i][6]
      };

      if (entry.type === 'schedule') {
        schedules.push(entry);
      } else if (entry.type === 'recommend') {
        recommendations.push(entry);
      }
    }

    // Sort by createdAt descending (newest first)
    schedules.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    recommendations.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    return {
      success: true,
      schedules: schedules,
      recommendations: recommendations,
      totalCount: schedules.length + recommendations.length
    };
  }

  // Save AI analysis result for a customer
  function saveCustomerPlant(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanPhone = String(data.phone || '').replace(/\D/g, '').slice(-10);

    if (!cleanPhone || cleanPhone.length !== 10) {
      return { success: false, error: 'Valid phone number required' };
    }

    if (!data.type || !['recommend', 'schedule'].includes(data.type)) {
      return { success: false, error: 'Type must be "recommend" or "schedule"' };
    }

    // Get or create UserPlants sheet
    let plantsSheet = ss.getSheetByName(TABS.USER_PLANTS);
    if (!plantsSheet) {
      plantsSheet = ss.insertSheet(TABS.USER_PLANTS);
      plantsSheet.appendRow(['ID', 'Phone', 'Type', 'ImageURLs', 'Data', 'PlantCount', 'CreatedAt']);
      plantsSheet.setFrozenRows(1);
    }

    const id = 'PLT' + Date.now().toString(36).toUpperCase();
    const imageURLs = Array.isArray(data.imageURLs) ? data.imageURLs.join(',') : (data.imageURLs || '');
    const plantData = typeof data.data === 'string' ? data.data : JSON.stringify(data.data || {});
    const plantCount = data.plantCount || 0;

    plantsSheet.appendRow([
      id,
      cleanPhone,
      data.type,
      imageURLs,
      plantData,
      plantCount,
      new Date()
    ]);

    logInfo('PLANT_SAVED', 'Customer plant analysis saved', {
      id: id,
      phone: cleanPhone,
      type: data.type,
      plantCount: plantCount
    });

    return {
      success: true,
      id: id,
      message: 'Plant analysis saved to your account'
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
  // DAILY SUMMARY EMAIL - Tomorrow's Bookings
  // ========================================

  function sendTomorrowBookingsSummary() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingsSheet = ss.getSheetByName(TABS.BOOKINGS);
    const bookingsData = bookingsSheet.getDataRange().getValues();

    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const tomorrowFormatted = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'EEE, dd MMM yyyy');

    const tomorrowBookings = [];

    for (let i = 1; i < bookingsData.length; i++) {
      const bookingDate = bookingsData[i][6];
      const bookingDateStr = parseDateToFormatted(bookingDate);

      if (bookingDateStr === tomorrowStr) {
        tomorrowBookings.push({
          bookingID: bookingsData[i][0],
          customerName: bookingsData[i][1],
          phone: bookingsData[i][2],
          gardenerName: bookingsData[i][5],
          timeSlot: formatTimeSlotForDisplay(bookingsData[i][7]),
          plantCount: bookingsData[i][8],
          address: bookingsData[i][9]
        });
      }
    }

    // Sort by time
    tomorrowBookings.sort((a, b) => {
      return timeToMinutes(a.timeSlot) - timeToMinutes(b.timeSlot);
    });

    // Build email
    let subject, body;

    if (tomorrowBookings.length === 0) {
      subject = '📅 PotPot: No bookings for ' + tomorrowFormatted;
      body = 'No bookings scheduled for tomorrow (' + tomorrowFormatted + ').';
    } else {
      subject = '📅 PotPot: ' + tomorrowBookings.length + ' booking(s) for ' + tomorrowFormatted;

      body = 'BOOKINGS FOR TOMORROW: ' + tomorrowFormatted + '\n';
      body += '='.repeat(50) + '\n\n';
      body += 'Total: ' + tomorrowBookings.length + ' booking(s)\n\n';

      for (let i = 0; i < tomorrowBookings.length; i++) {
        const b = tomorrowBookings[i];
        body += (i + 1) + '. ' + b.timeSlot + ' - ' + b.customerName + '\n';
        body += '   Phone: ' + b.phone + '\n';
        body += '   Plants: ' + b.plantCount + '\n';
        body += '   Gardener: ' + b.gardenerName + '\n';
        body += '   Address: ' + b.address + '\n';
        body += '   Booking ID: ' + b.bookingID + '\n';
        body += '\n';
      }
    }

    try {
      MailApp.sendEmail({
        to: 'potpot@atlasventuresonline.com',
        subject: subject,
        body: body
      });
      Logger.log('✅ Daily summary email sent: ' + tomorrowBookings.length + ' bookings');
    } catch (error) {
      Logger.log('❌ Failed to send daily summary: ' + error.toString());
    }

    return { success: true, bookingsCount: tomorrowBookings.length };
  }

  // ========================================
  // TRIGGER SETUP (Run once after updating)
  // ========================================

  function setupDailySummaryTrigger() {
    // Remove existing triggers for this function
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'sendTomorrowBookingsSummary') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }

    // Create new trigger at 6 PM IST
    ScriptApp.newTrigger('sendTomorrowBookingsSummary')
      .timeBased()
      .atHour(18)  // 6 PM
      .everyDays(1)
      .inTimezone('Asia/Kolkata')
      .create();

    Logger.log('✅ Daily summary trigger set! sendTomorrowBookingsSummary will run every day at 6 PM IST');
  }

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
