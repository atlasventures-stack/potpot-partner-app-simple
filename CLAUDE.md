# Claude Code Instructions - PotPot Partner App & Backend

## ONLY THESE 3 LIVE WEBSITES EXIST
1. https://www.potpot.online - Main website
2. https://potpot-booking-form.vercel.app - Booking form direct
3. https://potpot-partner-app-simple.vercel.app - Partner app

**Any other URLs are NOT in use. Ignore them.**

## ALWAYS DO FIRST
Read these files before any work:
1. `/Users/apple/Projects/potpot-booking-form/CLAUDE_CONTEXT.md`
2. `/Users/apple/Projects/potpot-partner-app-simple/CLAUDE_CONTEXT.md`

## This Folder Contains
- `index.html` - Partner/Gardener app UI
- `Code.gs` - Google Apps Script backend (shared by ALL 3 websites)

## Deploy Partner App
```bash
cd /Users/apple/Projects/potpot-partner-app-simple && vercel --prod
```

## Deploy Backend (Code.gs)
1. Open Google Apps Script editor
2. Replace Code.gs content
3. Deploy → Manage deployments → Create new deployment
4. Copy new URL and update API_URL in BOTH files:
   - `/Users/apple/Projects/potpot-booking-form/booking.html`
   - `/Users/apple/Projects/potpot-partner-app-simple/index.html`
5. Deploy both frontends with `vercel --prod`

## Critical Rule
Marketing ADs are running - verify everything before deploying.
