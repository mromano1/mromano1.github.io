# Google Drive setup

The backend now creates this structure automatically inside the Drive folder whose ID is in `Code.gs`:

- `Submissions/`
  - `Application Data/` — JSON records plus the JSON index used for Save/Load.
  - `<Project Title> - <Student or Team>/`
    - `<Project Title> - <Student or Team> - Project Plan` (Google Doc)
    - `<Project Title> - <Student or Team> - Project Schedule` (Google Sheet)

You do **not** need to create these folders yourself.

## Update the Apps Script
1. Open your existing Apps Script project.
2. Replace its `Code.gs` with the new `google-apps-script/Code.gs` in this package.
3. Confirm `DRIVE_FOLDER_ID` contains only the folder ID, not the full Drive URL.
4. Save.
5. Deploy > Manage deployments > Edit (pencil) > choose **New version** > Deploy.
6. Keep the same `/exec` Web App URL in the HTML pages if you updated the existing deployment.

The JSON remains visible in `Submissions/Application Data`; it is the editable data source that lets the website reload saved work. The readable Project Plan is updated as a Google Doc and the Project Schedule is updated as a Google Sheet rather than creating a new readable file on every save.
