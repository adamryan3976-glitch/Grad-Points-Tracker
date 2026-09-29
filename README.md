# WPS Grad Point Tracker

A web app for teachers to award and track student Grad Letter points. Students must meet the target in all four categories to earn the Letter:

| Category | Points required |
|---|---|
| Arts | 4 |
| Academics | 5 |
| Athletics | 8 |
| Service | 8 |

**How it's built**

- **GitHub Pages** hosts the app page (`index.html`, `config.js`).
- **Google Sheet + Apps Script** (`apps-script/Code.gs`) stores the classes, students and every point entry, and is the only place student data lives. GitHub holds code only, never student data.
- **Staff passcode:** teachers sign in with their name and a shared staff passcode you set in the Sheet. Every entry records who entered it.

---

## Part 1: Set up the Google Sheet (about 10 minutes)

1. Create a new Google Sheet named **WPS Grad Point Tracker**.
2. Open **Extensions > Apps Script**. Replace everything in `Code.gs` with the contents of [`apps-script/Code.gs`](apps-script/Code.gs). (No HTML file is needed in Apps Script.)
3. Save. Pick **setup** in the function dropdown, click **Run**, and approve the permission prompt.
4. Reload the Sheet. A **Grad Points** menu appears. Choose **Grad Points > Set staff passcode** and enter a passcode (6+ characters).
5. Fill in the tabs:

| Tab | What goes there |
|---|---|
| **Students** | `First Name`, `Last Name`, **`Grade`**, `Class` (e.g. Room 21). Only students with Grade **7** or **8** appear in the app. You can paste the whole school; other grades are simply ignored. Leave **Student ID** blank; the app fills it in (S0001…). Set **Active** to `No` for students who leave. |
| **Categories** | The four categories and their targets. Edit here if targets change. |
| **Activities** | Optional suggestions that pop up as teachers type. Teachers can type any activity. |
| **Points** | Every award: date, student, category, activity, points, note, who entered it. Don't edit by hand. |
| **Summary** | Auto-built totals for every Grade 7–8 student, sorted by last name, with **EARNED / In progress**. |

   Delete the two sample students when you add real ones.

   > **Already set up an earlier version?** Paste in the new `Code.gs`, then choose **Grad Points > Set up / repair sheets**. This adds the **Grade** column to your Students tab without touching your data. Fill in each student's grade, then redeploy (see "Making changes later"). The old **Classes** tab is no longer used and can be deleted.

## Part 2: Publish the data service

1. In Apps Script: **Deploy > New deployment**. Click the gear icon and choose **Web app**.
2. **Execute as:** Me. **Who has access:** **Anyone**.
   The GitHub Pages site can only reach the Sheet this way. The passcode is what keeps the data private: without it, the service returns nothing.
3. Click **Deploy** and copy the **Web app URL** (ends in `/exec`).

> If your school board account doesn't offer "Anyone", the board has blocked outside access to Apps Script. In that case the GitHub Pages version can't work with that account. Check with your board about where student records may be stored before using a personal Google account.

## Part 3: Publish on GitHub Pages

1. Put these files in your repo (root of the `main` branch):
   ```
   index.html
   config.js
   manifest.webmanifest      ← home-screen app settings
   favicon.ico
   icons/                    ← logo, Chrome tab icon, iPhone/Android home-screen icons
   apps-script/Code.gs
   README.md
   ```
2. Edit **`config.js`** and paste your Web app URL:
   ```js
   window.WPS_CONFIG = {
     API_URL: 'https://script.google.com/macros/s/AKfy.../exec'
   };
   ```
3. Commit and push.
4. On GitHub: **Settings > Pages > Build and deployment > Source: Deploy from a branch**, branch **main**, folder **/ (root)**. Save.
5. After a minute your site is live at `https://<your-username>.github.io/<repo-name>/`. Share that link and the passcode with teachers.

## Using the app

- Sign in with your name and the staff passcode. Tick "Keep me signed in" on your own device only.
- All Grade 7–8 students are listed alphabetically by last name, with their class and grade under each name. Each student shows progress bars for all four categories and a **★ Letter earned** badge once every target is met.
- To award a club or team: search for each student and tick them. Search matches first name, last name or class, so `Maya 21` finds Maya in Room 21. Ticks stay while you search, so you can build the whole team up one student at a time. The **Class** and **Grade** filters and **Select all shown** help with whole groups.
- Pick the category, **type the activity**, set the points (default 1), add a note, and press **Award**. Every ticked student gets the points in one step.
- Click a student's name to see their full history. **Remove** marks an entry as removed: it stays in the Sheet for the record but no longer counts.

## Making changes later

- **Classes, students, categories, suggestions:** edit the Sheet. Teachers see changes the next time they load the app.
- **Change the passcode:** Grad Points > Set staff passcode. Everyone is signed out and needs the new one.
- **App page (`index.html`):** commit and push; GitHub Pages updates in a minute or two.
- **Apps Script code:** after editing, go to **Deploy > Manage deployments**, click the pencil icon, choose **Version: New version**, and click **Deploy**. The URL stays the same, so `config.js` doesn't change.

## Notes

- Points follow the student's ID, so if a student changes classes or moves from Grade 7 to 8 their points come with them. Just update the Grade/Class cells.
- Grades can be entered as `7`, `Gr 7` or `Grade 7`. To change which grades are eligible, edit `ELIGIBLE_GRADES` at the top of `Code.gs`.
- After 25 wrong passcode attempts, the service locks for 10 minutes.
- The Apps Script URL in `config.js` is visible to anyone who views the site's code. That's expected; the passcode protects the data.

## Logo and app icons

The `icons/` folder was generated from the Winchester P.S. Grad Points Tracker logo:

| File | Used for |
|---|---|
| `icons/logo.png` | Logo in the app header and on the sign-in screen |
| `favicon.ico`, `icons/favicon-32.png` | Chrome tab icon (knight's helmet and cap, so it's readable at tiny sizes) |
| `icons/apple-touch-icon.png` | iPhone/iPad home-screen icon (Safari > Share > Add to Home Screen) |
| `icons/icon-192.png`, `icons/icon-512.png`, `icons/icon-maskable-512.png` | Android home-screen icon (Chrome > ⋮ > Add to Home screen / Install app) |

Once it's added to a home screen, the app opens full-screen like a regular app, labelled **Grad Points**. Phones cache icons, so if the icon changes later, remove the shortcut and add it again.
