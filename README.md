# Family create-only reminder page

This package adds a separate public submission page and an authenticated family-reminder admin page. Your existing jobs.html and its backend are unchanged.

## What works

- Family members create reminders with their own email/mobile, job roles, locations, work modes, shifts, times, time zone, job age and custom weekdays.
- They cannot list existing entries, delete them or change their enabled status. The server checks the admin token on every privileged operation; editing the HTML does not bypass this.
- Admin can list, enable/disable and delete family reminders.
- New preferences are stored in your private Google spreadsheet. Public callers receive only their newly created reference, never other records.

## Setup

1. Create a NEW project at https://script.google.com/ . Do not overwrite your current reminder script.
2. Paste Code.gs into the script editor.
3. Add HTML files named Family and Admin; paste the matching HTML contents.
4. In Project Settings → Script properties, add FAMILY_ADMIN_TOKEN. Use a unique random password of at least 32 characters (a password manager can generate it). Keep it private. No admin token is embedded in Family.html.
5. Run setupFamilyStorage_ from the editor once and authorize it. This creates the spreadsheet and stores its ID in FAMILY_SHEET_ID. Do not share that spreadsheet with family members.
6. Deploy → New deployment → Web app. Set Execute as: Me and Who has access: Anyone. Authorize the deployment as needed. This grants public access to the create-only endpoint, not to your spreadsheet or administration data. Account policies may restrict this option.
7. Copy the deployed URL ending /exec.

Share with family:
  https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec

Admin page (for you):
  https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?page=admin

Anyone can open the empty admin login page, but listing or changing entries requires FAMILY_ADMIN_TOKEN. Keep the token out of URLs, source files and family messages. Admin HTML does not store it in localStorage.

Deploy a new version after code changes; keep using the /exec URL rather than the editor /dev URL.

## Reminder delivery integration

The package registers and administers preferences. It does not send email or WhatsApp and does not silently attach itself to your unknown existing sender. The family page clearly states that delivery starts after the admin connects the sender.

getFamilyRemindersForDelivery_() is a PRIVATE helper for sender code in this same project. It returns enabled records and their settings. A sender must honor repeat/repeatDays, times/timezone, individual mode/roles, locations, posted-within days and workModes/shifts. WhatsApp needs your supported messaging provider.

To integrate with your existing scheduler, supply its source so these records can be read through a server-side integration and delivery can be tested. Do not make this helper public or expose the admin token to clients. Never connect the family form to the old unrestricted save action.

## Limits

- One recipient per submission. Family can create another reminder for another address.
- No family edit/delete links, no family access to stored records.
- Anyone holding the public URL can submit preferences. This is create-only access, not verified family identities. For named family-only admission, add sign-in or individual invitation validation.
- Registration validates fields, ignores attempted injected IDs/actions/enabled status, rejects exact duplicate settings, uses a write lock and caps entries at 1,000. Real email/mobile ownership is not verified.
- Existing jobs.html and existing backend permissions are unchanged. Share ONLY the new family URL. It cannot grant access to the original admin page, but it does not fix any pre-existing weaknesses in that original backend.
- Apps Script deployment and external account authorization have not been performed in this workspace. Browser/end-to-end deployment and actual delivery remain untested.

Official deployment reference: https://developers.google.com/apps-script/guides/web
Private server function convention: https://developers.google.com/apps-script/guides/html/communication#private_functions

## Included files

- Code.gs — backend with permissions enforced server-side.
- Family.html — create-only family page (Apps Script HTML service).
- Admin.html — authenticated administration for the family records.
- tests.cjs — local permission/validation regression checks using Apps Script mocks.
- README.md — setup and integration instructions.
