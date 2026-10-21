# New communications and risk forms

Upload `communications-plan-form.html`, `risk-management-form.html`, and the updated `project-management.html` and `assets/style.css` to GitHub Pages.

In each new HTML file, replace `PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE` with the same deployed `/exec` URL used by your existing forms.

Replace the code in your existing Apps Script project with `google-apps-script/Code.gs`. Then choose Deploy → Manage deployments → Edit → New version → Deploy. Keep the existing deployment so its URL does not change.

The backend automatically creates/updates readable Google Docs directly in `Submissions` and JSON records in `Submissions/Application Data`. Each form uses the student's email address and form type to reload the draft.

**Important:** The supplied Word document explicitly provides the risk tables but only references the Communications Plan as a separate deliverable and shows a contact list. The communications matrix, approach, and escalation sections are proposed fields, not transcribed from the attachment.

Before use, test each new form with a sample email, save, load, and final submission. No live Google deployment was performed here. Email-only lookup does not authenticate students; consider Google sign-in or email verification before collecting sensitive information.
