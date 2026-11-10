# Northstar Academy admissions

A responsive single-page admission app with a Google Apps Script API, Google Sheets records, Drive photo storage, and Razorpay checkout. Open `index.html` directly to explore the clearly labeled local demo, or deploy the Apps Script files for live use.

## Google Apps Script setup

1. Create a new Apps Script project and add `Code.gs` and `index.html` from this folder (the HTML file must be named `index` in the Apps Script editor).
2. Run `initializeApp()` once and authorize Sheets, Drive, and external requests. It creates the spreadsheet tabs and photo folder. Copy the spreadsheet URL from the execution log.
3. Deploy the Web App, sign in as an admin, open **Razorpay setup**, and import the Razorpay test API keys CSV. The key ID and secret are saved to Apps Script Script Properties (`RAZORPAY_KEY_ID` and `RAZORPAY_KEY_SECRET`); the secret is never written to the HTML or source files. Alternatively, enter those two properties manually in **Project Settings → Script properties**.
4. In the editor, run `setInitialAdmin('admin@example.com', 'a-long-unique-password')` once. Use a unique password of at least 12 characters. Do not save a real password in project source.
5. Deploy as a **Web app**, execute as your account, and grant access to anyone who can use the public admission form. The UI picks up the deployed URL automatically when served by Apps Script.

The app uses a POST JSON API (`action` values in `doPost`) through Fetch. If you host the HTML elsewhere, set `localStorage.setItem('northstar_api_url', 'YOUR_DEPLOYED_WEB_APP_URL')` in that site's browser before loading the app. Apps Script deployment and browser cross-origin behavior can vary by account; use the deployed Web App URL and confirm that POST requests return JSON in your target browser.

## Included flows

- Public form with required field validation, photo preview and 2 MB image limit, class fee display, and Razorpay checkout.
- Razorpay order creation on the server and signature plus captured-payment verification before creating a student row.
- Admin-only Razorpay CSV import for test credentials, stored in Apps Script properties.
- Optional UPI QR and bank transfer details managed from the admin menu. Manual transfers create registrations marked `Pending verification`; an admin confirms receipt before the student is marked paid.
- Unique admission form number, class roll number, Drive photo link, and printable registration slip.
- Admin login, class and fee metrics, search/filter, student detail view, CSV export, editable fee schedule, class/batch catalog management, and admin user/password management.
- Admin institute profile settings, class-specific roll number prefixes and next-number settings, and student edit/delete actions. Deletes are moved to `Deleted_Students` for recovery.
- Optional 12-digit Aadhaar field stored with the student record and included in the admin detail view/export.
- Local demo mode when the page is opened as a static file. Demo records stay in that browser's local storage; demo payments are not real charges.

## Security and operations notes

- `Admin_Credentials` stores a SHA-256 password hash; sessions are short-lived in Apps Script Cache. Require HTTPS and restrict who can access the admin URL in production.
- Run `initializeApp()` again after replacing the script files; it adds the Aadhaar, batch and roll-prefix columns plus the `Deleted_Students` archive tab while preserving active student rows. Redeploy the Web App to publish the updated version.
- Aadhaar numbers are sensitive personal data and are stored in the Sheets row and admin CSV export. Limit spreadsheet and admin access to staff who need it, and follow your organization's retention and privacy requirements.
- Manual QR/bank transfers are not automatically verified by the app. Admins should confirm the credit in the institute's UPI/bank account before using **Confirm payment** in the dashboard.
- After deploying, admins can set the institute UPI ID, payee, bank name, account holder, account number, IFSC and branch under **UPI / bank details**. Students see a fee-specific UPI QR and configured bank instructions during review.
- Drive photos are created with anyone-with-link view access so the printed slip can display the image. For stricter privacy, replace that setting with authenticated image delivery.
- Apps Script quotas, spreadsheet concurrency, payment provider account setup, and deployment access policy apply. The script uses a lock for roll assignment and checks Razorpay's captured payment amount against the current configured fee.
- PhonePe is not wired in this implementation; Razorpay sandbox and production keys are supported through Script Properties.
