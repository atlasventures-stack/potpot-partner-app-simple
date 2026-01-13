# PotPot Backend & Partner App Context for Claude

**Last Updated:** 2026-01-13

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
https://script.google.com/macros/s/AKfycbxVIGgzKaltYhxFoWWkJJSEln6bYQa7QHNAtGzyvgbdYmuzDuN1zzC7IlN2RpQ4WNnR2w/exec
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
- `?action=checkPaymentStatus&paymentLinkId=X` - Check Razorpay payment

### POST Requests
- Booking creation (default) - Creates new booking
- `{action: 'saveServiceReport', ...}` - Save service completion report
- `{action: 'createPaymentLink', ...}` - Create Razorpay payment link
- `{action: 'updatePaymentStatus', ...}` - Update payment status

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
- 1-20 plants: 60
- 21-35 plants: 90
- 36-50 plants: 120
- 51+ plants: 180

## Google Sheets Tabs
- `Bookings` - All bookings
- `Availability` - Slot availability tracking
- `GardenerZones` - Gardener assignments by pincode
- `ServiceReports` - Completed service reports
- `Admins` - Admin phone numbers for dashboard access
- `CustomerData` - Customer info for follow-ups

## WATI Templates
- `booking_confirmation` - Sent on new booking
- `pre_service_com_time` - Day-before reminder
- `post_service_checkup` - 5-day follow-up
- `nps_form` - Sent after service completion

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
  - 1-20 plants → ₹399
  - 21-35 plants → ₹599
  - 36-50 plants → ₹799
  - 51-75 plants → ₹999
  - 76-100 plants → ₹1199
  - 101+ plants → ₹1499
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
