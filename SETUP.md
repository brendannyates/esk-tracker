# ESK Tracker — setup (about 10 minutes, once)

## 1. Host the files (GitHub Pages, free)
1. Sign in at github.com → **New repository** → name it `esk-tracker` → Public → Create.
2. **Add file → Upload files** → drag in everything from this folder (taskpane.html, commands.html, index.html, demo.html, and the `assets` folder) → Commit.
3. **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / root → Save.
4. After a minute, open `https://<your-github-name>.github.io/esk-tracker/taskpane.html`.

The site only serves the app's code. Patient data stays in the workbook.
(If IT prefers an internal host such as Azure Static Web Apps, upload the same files there. Any HTTPS address works.)

## 2. Get the manifest
On that page, click **Download manifest**. It fills in its own web address automatically.

## 3. Add it to the workbook
Open the tracker in Excel on the web → **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in** → choose `esk-tracker-manifest.xml`.
An **ESK Tracker** button appears on the Home tab.

## 4. First run
After it reads the workbook, the tracker asks **Enter your name** once a day on each computer (tech, RN, MA or provider), with the people registered in this workbook listed below. Pick your name, or choose **I'm not on the list** to add yourself. Nothing about you carries over to another clinic's workbook: each workbook has its own list, and each computer remembers who used it per workbook.

The clinic's settings (clinic code, location, practice manager, order and cc emails, MADRS cc) belong to the workbook. The clinic and location start from the workbook's name when it says (for example "Castro ESK Tracker"); otherwise they're blank. The first person to fill in a blank one in Settings saves it for everyone; after that, changing it needs Admin (the workbook's PIN or the MHS admin code).

Emails end with "Thank you," and no signature, because Outlook adds yours. If it doesn't, open Settings › **Add a signature** and type it or pick your Outlook signature file (`%APPDATA%\Microsoft\Signatures`, the .txt or .htm).

Optional: tick "Open this panel automatically whenever this workbook opens".

## Other staff
Each person uploads the same manifest once. Your M365 admin can instead deploy it to the team from the admin center.

## Updating
Replace the files on the host. Everyone gets the new version the next time the panel opens. No new manifest is needed.

## Try it without Excel
Open `demo.html`. It uses fake data only.
