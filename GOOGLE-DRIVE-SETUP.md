# Connect the forms to Google Drive

1. In Google Drive, create a folder for student submissions. Copy the folder ID from its URL.
2. Go to script.google.com and create a new Apps Script project.
3. Replace the default code with `google-apps-script/Code.gs`.
4. Replace `PASTE_GOOGLE_DRIVE_FOLDER_ID_HERE` with your folder ID.
5. Deploy > New deployment > Web app. Execute as **Me**. Choose the access setting appropriate for your students/account policy. Authorize the script and deploy it.
6. Copy the Web App URL ending in `/exec`.
7. In BOTH `project-plan-form.html` and `project-schedule-form.html`, replace `PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE` with that URL.
8. Commit/push the files to GitHub.

## How saving works
Each student/team receives a browser-stored record ID after the first save. Save Draft updates the same Drive JSON record. Submit Final marks that saved record as final. Load Saved Work retrieves it. The Drive folder also contains an index file used by the script.

## Important
GitHub Pages cannot keep a Google secret safely. The Apps Script Web App is the server-side bridge. Do not put private API keys in the HTML. Test with a non-sensitive sample before student use, and use your organization's approved Google access/sharing policy.
