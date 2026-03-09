# SecondBrain Add-in POC — Testing Guide

## What you're testing

One question: **does `body.setAsync()` inject text into the compose window
when you click Reply in Classic Outlook desktop?**

---

## Files

```
addin-poc/
  manifest.xml     <- Tells Outlook what the add-in does (UPDATE YOUR URL)
  commands.html    <- Silent background host page
  addin.js         <- The event handler (~100 lines)
  icon-64.png      <- Placeholder (any 64x64 PNG)
  icon-128.png     <- Placeholder (any 128x128 PNG)
```

---

## Step 1: Host the files over HTTPS (10 min)

### Option A: GitHub Pages (easiest for POC)

1. github.com -> New repository -> `secondbrain-addin-poc` -> Public
2. Upload all files from this folder
3. Settings -> Pages -> Source: Deploy from branch -> main -> /root -> Save
4. Wait 2 minutes. URL: `https://YOUR-USERNAME.github.io/secondbrain-addin-poc`

### Option B: Azure Static Web Apps

1. Azure Portal -> Create Resource -> Static Web App -> Free tier
2. Connect to GitHub repo with these files
3. URL: `https://YOUR-APP.azurestaticapps.net`

---

## Step 2: Update manifest.xml with your URL (2 min)

Find-and-replace all instances of `YOUR-HOSTING-URL` with your actual URL
(no trailing slash). There are 4 places:

- IconUrl
- HighResolutionIconUrl
- AppDomain
- Commands.Url

Example:
```
https://YOUR-HOSTING-URL  ->  https://jameslarkey.github.io/secondbrain-addin-poc
```

Re-upload the updated manifest.xml to GitHub.

---

## Step 3: Create placeholder icons (2 min)

Create any two PNG files named `icon-64.png` (64x64px) and `icon-128.png`
(128x128px). Upload to the same repo. Any image works for testing.

---

## Step 4: Deploy the add-in (5 min)

### Quick way: Sideload in Outlook Web (instant, no admin needed)

1. Go to outlook.office.com
2. Open any email -> click ... menu -> Get Add-ins
3. My add-ins -> Custom add-ins -> + Add from URL
4. Paste: `https://YOUR-USERNAME.github.io/secondbrain-addin-poc/manifest.xml`
5. Install

This syncs to Classic Outlook desktop within ~5 minutes.

### Admin way: M365 Admin Center (installs for specific users)

1. admin.microsoft.com -> Settings -> Integrated Apps -> Upload custom apps
2. Upload your manifest.xml
3. Assign to yourself (and optionally one broker)
4. Takes 5-10 minutes to propagate

---

## Step 5: Run the tests (10 min)

Send yourself 2-3 test emails from a different account first.

### Test 1: Inline Reply (CRITICAL — this is Max's workflow)
1. Open Classic Outlook desktop
2. Select a test email in reading pane
3. Click Reply (inline compose appears at bottom of reading pane)
4. WATCH: Does the draft text appear with yellow "SecondBrain POC test draft"?

### Test 2: Pop-out Reply
1. Select an email
2. Double-click to open, then click Reply (or Shift+Reply)
3. WATCH: Does the draft text appear?

### Test 3: Outlook Web Reply
1. Go to outlook.office.com
2. Click Reply on a test email
3. WATCH: Does the draft text appear?

### Test 4: New Compose (should NOT inject)
1. Click New Email
2. WATCH: Compose should open blank — no injection
3. Confirms the RE:/FW: filtering works

---

## Step 6: Report results

| Test | Result | Notes |
|------|--------|-------|
| Inline reply (reading pane) | Works / Fails / Partial | |
| Pop-out reply | Works / Fails / Partial | |
| Outlook Web reply | Works / Fails / Partial | |
| New compose (no injection) | Correctly blank / Injected | |
| Console errors (if any) | | |
| Outlook version | | File -> Office Account -> About Outlook |

**The inline reading pane result is the critical one.**

---

## Debugging

### Outlook Web: F12 -> Console -> filter by "SecondBrain"

### Classic Outlook Desktop:
- Ctrl+Shift+I (may work on some versions)
- Or: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-f12-tools-ie

### Expected console output (success):
```
[SecondBrain] Office.js ready. Host: Outlook Platform: PC
[SecondBrain] Compose window opened
[SecondBrain] Subject value: RE: Test email subject
[SecondBrain] Reply detected! Attempting draft injection...
[SecondBrain] Conversation ID: AAQkAGI...
[SecondBrain] Draft injected successfully!
```

### Expected console output (new email, correctly skipped):
```
[SecondBrain] Compose window opened
[SecondBrain] Subject value:
[SecondBrain] New compose (not a reply) - skipping injection
```

---

## What happens next

**If inline reply works**: Green light for Phase 1. Your existing Python
backend (voice learning, classification, draft generation) stays intact.
You add a `/drafts/{conversationId}` endpoint and update addin.js to call
it instead of using hardcoded text. ~1-2 days of work.

**If inline fails but pop-out works**: Still viable. Add a notification
banner in reading pane: "Draft ready -> click to reply". Opens pop-out
compose with draft injected. One extra click, still saves massive time.

**If both fail**: body.setAsync() is blocked in your environment. Pivot to
taskpane approach (side panel shows draft, one-click copy to clipboard).
