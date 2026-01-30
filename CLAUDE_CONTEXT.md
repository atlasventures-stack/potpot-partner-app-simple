# PotPot Backend & Partner App Context for Claude

**Last Updated:** 2026-01-26

## ONLY THESE 3 LIVE WEBSITES EXIST
1. https://www.potpot.online - Main website
2. https://potpot-booking-form.vercel.app - Booking form direct
3. https://potpot-partner-app-simple.vercel.app - Partner app

**Any other URLs are NOT in use. Ignore them.**

## This Repository Contains
- `index.html` - Partner/Gardener app UI
- `Code.gs` - Google Apps Script backend (shared by ALL 3 websites)

## Current Backend API
```
https://script.google.com/macros/s/AKfycbzKY6z_05v5Gp7S_7znsyGiee6ySJu8iqY6P8CjDvKzflPI_pNlacduX50-mXhRNL5DxA/exec
```

## How to Deploy Backend (Code.gs)
1. Copy contents of `Code.gs`
2. Go to Google Apps Script editor
3. Paste and save
4. Deploy → New Deployment → Web App
5. Copy the new URL and update API_URL in BOTH:
   - `/Users/apple/Projects/potpot-booking-form/booking.html`
   - `/Users/apple/Projects/potpot-partner-app-simple/index.html`
6. Deploy both frontends: `vercel --prod` in each folder

## API Endpoints

### GET Requests
- `?action=validatePincode&pincode=X` - Check if pincode is serviceable
- `?pincode=X&lat=Y&lng=Z&plantCount=N` - Get available slots
- `?action=getGardenerJobs&phone=X` - Get gardener's jobs for today
- `?action=getAdminReports&phone=X` - Get admin dashboard data
- `?action=checkPaymentStatus&paymentLinkId=X` - Check Razorpay payment (DEPRECATED - not used by partner app)

### POST Requests
- Booking creation (default) - Creates new booking
- `{action: 'saveServiceReport', ...}` - Save service completion report
- `{action: 'createPaymentLink', ...}` - Create Razorpay payment link (DEPRECATED - not used by partner app)
- `{action: 'updatePaymentStatus', ...}` - Update payment status (used when gardener marks payment complete)

## Slot System Config
```javascript
const SLOT_CONFIG = {
  START_HOUR: 8, START_MIN: 30,
  END_HOUR: 18, END_MIN: 0,
  SLOT_INTERVAL: 30,
  LUNCH_START: 13, LUNCH_END: 14,
  HARD_CUTOFF: 18.5,
  TRAVEL_BUFFER: 30
};
```

## Service Duration (minutes)
- 0 plants (old bookings): 90
- 0-10 plants: 60 (no travel buffer)
- 11-30 plants: 60
- 31-50 plants: 120
- 51+ plants: 180

## Pricing Tiers
- 0-10 plants: ₹499
- 11-30 plants: ₹699
- 31-50 plants: ₹899
- 51-75 plants: ₹1199
- 76-100 plants: ₹1399
- 101-150 plants: ₹1499

## Google Sheets Tabs
- `Bookings` - All bookings
- `Availability` - Slot availability tracking
- `GardenerZones` - Gardener assignments by pincode
- `ServiceReports` - Completed service reports
- `Admins` - Admin phone numbers for dashboard access
- `CustomerData` - Customer info for follow-ups
- `Holidays` - Gardener leave/holiday dates

## WATI Templates
- `booking_confirmation` - Sent on new booking
- `pre_service_com_time` - Day-before reminder
- `post_service_checkup` - 5-day follow-up
- `nps_form` - Sent after service completion

## Recent Changes (2026-01-26)

### Simplified Payment Flow (Partner App)
- **Old flow (removed):** Service complete → Create Razorpay link → Generate dynamic QR → Poll for payment status → Auto-update on payment
- **New flow:** Service complete → Show static UPI QR → Gardener asks customer → Gardener taps "Payment Done" button → Update sheet + Amplitude event

#### What was removed:
- Razorpay `createPaymentLink` API call
- Razorpay `checkPaymentStatus` polling (every 5 seconds)
- Dynamic QR generation from payment link
- `paymentLinkId` and `paymentCheckInterval` variables

#### What was added:
- Static UPI QR code placeholder (replace `https://via.placeholder.com/200x200?text=UPI+QR` with actual QR)
- "Payment Done" button (`markPaymentComplete()` function)
- Manual confirmation flow - gardener confirms with customer verbally

#### Functions changed:
- `initPaymentSection()` - Simplified, only calculates amount and shows QR
- `markPaymentComplete()` - NEW - calls `updatePaymentStatus` API when gardener taps button
- `skipPayment()` - Simplified, just hides payment section
- `backToDashboard()` - Simplified, removed interval cleanup

### Amplitude Server-Side Tracking (Code.gs)
- **Problem:** Events happen in partner app (gardener's device), but should track on customer's Amplitude profile
- **Solution:** Async server-side Amplitude tracking via one-time triggers (doesn't block gardener)
- **API:** Amplitude HTTP API (`https://api2.amplitude.com/2/httpapi`)
- **User matching:** Uses customer phone (last 10 digits) as user_id - matches frontend booking.html

#### Functions Added:
- `scheduleAmplitudeEvent(eventData)` - Creates one-time trigger (5 sec delay), stores event data
- `processDelayedAmplitudeEvent(e)` - Triggered function that sends to Amplitude, cleans up
- `trackServiceCompleted(data)` - Called from `saveServiceReport()`
- `trackPaymentSuccess(reportRow)` - Called from `updatePaymentStatus()`

#### Event: Service Completed
- **Triggered from:** `saveServiceReport()` when gardener submits service report
- **Properties:** booking_id, gardener_id, gardener_name, total_plants, red_zone_plants, bugs_found, repotted_plants, service_started_at, service_completed_at, service_duration_minutes, amount

#### Event: Payment Success
- **Triggered from:** `updatePaymentStatus()` when payment status is "Paid"
- **Properties:** booking_id, gardener_id, gardener_name, total_plants, amount, currency (INR), payment_method (Razorpay)

### Partner App Bug Fix: customerPhone String Conversion
- **Bug:** `TypeError: job.customerPhone.replace is not a function`
- **Cause:** Google Sheets returns phone numbers as numbers, not strings
- **Fix:** Wrapped in `String()`: `String(job.customerPhone).replace(...)`
- **Location:** Partner app index.html, tel: link in job cards

## Recent Changes (2026-01-24)

### Lead Source Tracking (Code.gs)
- **New column:** `LeadSource` in Bookings sheet (Column AE - after column AD)
- **Values:** Facebook, Google, Organic, Direct
- **Stored in:** `createBooking()` function, added to bookingRow at index 30
- **Setup required:** Add "LeadSource" header in column AD of Bookings sheet

### Gardener Holidays Module
- **New sheet:** `Holidays` with columns: GardenerID, StartDate, EndDate, Reason, CreatedAt
- **How to add holiday:** Add row to Holidays sheet (e.g., G002, 25/01/2026, 27/01/2026, Sick leave)
- **Single day holiday:** Leave EndDate blank or same as StartDate
- **What happens:**
  1. Dates where gardener is on holiday show 0 available slots
  2. Load balancing considers holidays (gardener with more holidays is deprioritized)
  3. If ALL gardeners for a pincode are on holiday → that date shows no slots
- **Helper functions:** `isGardenerOnHoliday()`, `countHolidayDays()`
- **Response includes:** `holiday: true` flag on dates where gardener is on holiday

### Gardener Load Balancing
- **Problem:** Multiple gardeners can serve the same pincodes (G001/G004 share pincodes, G002/G003 share pincodes), but old code always picked the FIRST matching gardener
- **Solution:** Load balancing based on bookings + holidays count in `getAvailableSlots()` function
- **How it works:**
  1. When customer enters pincode, find ALL gardeners serving that pincode
  2. If only 1 gardener → use them (same as before)
  3. If multiple gardeners → count each gardener's bookings + holidays in the next 7 days
  4. Assign the gardener with fewer blocked days
- **Logging:** `LOAD_BALANCE` entries in Logs sheet show booking counts, holiday counts, and selection decision
- **Location:** `getAvailableSlots()` function

## Recent Changes (2026-01-13)

### Partner App: Bugs Found Toggle
- Added "🐛 Bugs found in plants (पौधों में कीड़े मिले)" toggle checkbox in Complete Service step
- Partner can mark if bugs/pests were found during service
- `bugsFound` field sent to backend as "Yes" or "No"
- Shows in completion summary screen

### Partner App: Dynamic Payment Calculation
- **New payment logic based on Amount column in Bookings sheet:**
  - `Amount = 0` → No payment section shown
  - `Amount = blank/empty` → Calculate from partner's entered plant count
  - `Amount = value` → Show that exact amount
- **Pricing tiers (when calculating from plants):**
  - 0-10 plants → ₹499
  - 11-30 plants → ₹699
  - 31-50 plants → ₹899
  - 51-75 plants → ₹1199
  - 76-100 plants → ₹1399
  - 101-150 plants → ₹1499
- Uses `reportData.totalPlants` (actual plants serviced, not booked)

### Partner App: UI Fixes
- Fixed bugs toggle layout (checkbox and text now side-by-side)
- Added `flex-direction: row` and `flex-shrink: 0` for proper alignment

## Recent Changes (2026-01-11)

### NPS Delayed Sender (Service Report Speed Fix)
- **Problem:** Gardeners were waiting too long after submitting service reports (WATI NPS form was blocking)
- **Solution:** NPS forms sent via one-time delayed trigger (no recurring jobs)
- **How it works:**
  1. `saveServiceReport()` calls `scheduleNPSForm()` - creates one-time trigger, returns immediately
  2. Trigger fires after 2 minutes → `processDelayedNPS()` sends NPS → trigger auto-deletes
  3. No setup needed - each service report creates its own trigger automatically
- **Result:** Service report submission is instant, NPS arrives ~2 min later

### Daily Summary Email
- **Function:** `sendTomorrowBookingsSummary()` - sends email at 7 PM IST with next day's bookings
- **Trigger:** `setupDailySummaryTrigger()` - run ONCE to set up (already done ✅)
- **Email to:** potpot@atlasventuresonline.com

### Critical Bug Fix: Missing Bookings (Silent Save Failures)
- **Root Cause:** `appendRow()` in Google Apps Script can fail silently without throwing errors
- **Problem:** Booking ID was generated BEFORE save, so even if save failed, WhatsApp confirmation was sent
- **Fix 1:** Added `SpreadsheetApp.flush()` to force write completion
- **Fix 2:** Added 300ms delay for persistence
- **Fix 3:** Added verification check - reads sheet to confirm booking exists
- **Fix 4:** Added retry logic - if booking not found, tries one more time
- **Fix 5:** Added email backup to potpot@atlasventuresonline.com for EVERY booking
- **Email subject:** Shows ✅ if verified, ❌ if failed (check immediately)

## Recent Changes (2026-01-10)
### Critical Bug Fix: Duplicate Bookings
- **Root Cause:** Google Sheets returns Date objects for dates/times, but `instanceof Date` failed in Apps Script
- **Fix 1:** Added multiple Date object detection methods (`typeof dateValue.getTime === 'function'`)
- **Fix 2:** Handle time formats with seconds (`11:00:00 AM`) via updated regex
- **Fix 3:** Added `SLOT_TAKEN` validation in `createBooking()` as second layer of protection
- **Fix 4:** Added `testDuplicateDetection()` function to verify fix with actual data

### Other Fixes
- Fixed `timeToMinutes` to handle both "2:00 PM" and "14:00" formats
- Fixed `getServiceDuration` to handle string ranges like "0-20", "20-35"
- Fixed GardenerID comparisons to use String() conversion
- Fixed phone number comparisons to extract last 10 digits consistently
- Fixed sorting to use timeToMinutes() instead of localeCompare

### Partner App Fixes (2026-01-10)
- Added `formatTimeSlotForDisplay()` helper to convert Date objects to readable "8:30 AM" format
- Fixed `getGardenerJobs()` - now returns properly formatted time slots instead of Date object garbage
- Fixed `getAdminReports()` - same fix for admin dashboard
- Added job sorting in `getGardenerJobs()` - jobs now appear in chronological order (earliest first)
