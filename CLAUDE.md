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

## GitHub Repository
https://github.com/atlasventures-stack/potpot-partner-app-simple

## Deployment Workflow (MANDATORY)

**CRITICAL: ALWAYS use explicit `cd` before EVERY command. The shell persists directory state, so running `vercel --prod` without `cd` will deploy the WRONG project. NEVER assume working directory is correct.**

### For Partner App UI
```bash
cd /Users/apple/Projects/potpot-partner-app-simple && git add . && git commit -m "description of changes"
cd /Users/apple/Projects/potpot-partner-app-simple && git push
cd /Users/apple/Projects/potpot-partner-app-simple && vercel --prod
```

### For Backend (Code.gs)
1. Edit `Code.gs`
2. Open Google Apps Script editor
3. Replace Code.gs content
4. Deploy → Manage deployments → Create new deployment
5. Copy new URL and update API_URL in BOTH:
   - `/Users/apple/Projects/potpot-booking-form/booking.html`
   - `/Users/apple/Projects/potpot-partner-app-simple/index.html`
6. Commit & push both repos
7. `vercel --prod` for both folders

## After Every Deployment
Update `CLAUDE_CONTEXT.md` with:
- What was changed
- Date of deployment
- New API URL (if backend changed)

## Critical Rule
Marketing ADs are running - verify everything before deploying.
