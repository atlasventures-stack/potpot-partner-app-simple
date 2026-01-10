# PotPot Backend & Partner App Context for Claude

**Last Updated:** 2026-01-10

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
https://script.google.com/macros/s/AKfycbxHrc_rMELelwhFgkhO0rxqoapDkLMt-h2AYnZRBjVFfK-3rST5lDEcVe4-zJUbRkSa/exec
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
