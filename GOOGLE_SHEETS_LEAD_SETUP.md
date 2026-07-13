# Prebook Lead Capture Setup (No Backend)

## Week 1: Google Sheets + Instant Notifications

1. Create a Google Sheet named `Prebook Leads`.
2. Add header row:
   - `timestamp`
   - `name`
   - `phone`
   - `email`
   - `tripDates`
   - `pickupLocation`
   - `vehicleType`
   - `source`
3. Open `Extensions -> Apps Script`.
4. Paste this script into `Code.gs`:

```javascript
const SHEET_NAME = 'Prebook Leads';
const NOTIFY_EMAIL = 'premiumecorentals@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'timestamp',
        'name',
        'phone',
        'email',
        'tripDates',
        'pickupLocation',
        'vehicleType',
        'source'
      ]);
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.phone || '',
      data.email || '',
      data.tripDates || '',
      data.pickupLocation || '',
      data.vehicleType || '',
      data.source || ''
    ]);

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New Prebook Lead: ' + (data.name || 'Unknown'),
      body:
        'New lead received\n\n' +
        'Name: ' + (data.name || '') + '\n' +
        'Phone: ' + (data.phone || '') + '\n' +
        'Email: ' + (data.email || '') + '\n' +
        'Trip Dates: ' + (data.tripDates || '') + '\n' +
        'Pickup Location: ' + (data.pickupLocation || '') + '\n' +
        'Vehicle Type: ' + (data.vehicleType || '') + '\n' +
        'Source: ' + (data.source || '')
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

5. Deploy as web app:
   - `Deploy -> New deployment -> Web app`
   - Execute as: `Me`
   - Who has access: `Anyone`
6. Copy the web app URL.
7. Paste URL in `index.html` in `const sheetsEndpoint = "";`.
8. Test once from the site form and confirm:
   - New row appears in sheet
   - Notification email arrives

## Weeks 2-4: CRM Migration Plan (HubSpot/Zoho/Pipedrive)

1. Keep Google Sheet as backup for one week after CRM go-live.
2. Add CRM stages:
   - New Lead
   - Contacted
   - Quote Sent
   - Booked
   - Lost
3. Send website form directly to CRM webform/API.
4. Keep notification alerts for `New Lead` and `No response in 30 minutes`.
5. Add automation:
   - Auto-assign by pickup date urgency
   - Reminder task if no follow-up within 1 hour
6. Weekly QA:
   - Compare form submissions vs CRM records
   - Reconcile missing leads

## Current Site Wiring

The form is already live in `index.html` and emits GTM events:
- `lead_prebook_submit`
- `lead_prebook_submit_success`
- `lead_prebook_submit_error`
- `comeback_help_availability_click`

Update only the `sheetsEndpoint` value to activate Week 1 collection.
