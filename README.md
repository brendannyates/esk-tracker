# ESK Tracker

Excel add-in for the Esketamine Master Tracker workbook. It covers today's schedule, med pulls, administration, receiving, disposal, and Besse/Genoa ordering.

- App: https://brendannyates.github.io/esk-tracker/taskpane.html
- Demo (fake data): https://brendannyates.github.io/esk-tracker/demo.html
- Setup: see SETUP.md

This site serves only the app's code. Patient data stays in the workbook.

## Sohar benefits
In the Sohar Eligibility add-in, export the **Final tracker file** (CSV). In the ESK Tracker, select **⇪ Import Sohar benefits** at the top and choose that file. You'll see a preview first: how many patients match the Master Patient Tracker, who isn't eligible, and any ESK patients who aren't on the tracker yet (listed by MRN, not imported).

Importing writes to the **Benefits** tab (table `tblBenefits`), one row per patient, using the latest Sohar check for each. Importing again replaces those patients' rows and keeps everyone else's. The Master Patient Tracker isn't changed. Each patient's Details shows a one-line summary (copay, deductible and out-of-pocket left, date of the Sohar check), with a red flag if Sohar says they aren't eligible.

Source is in `src/`. Run `python3 build.py` to rebuild into `dist/`, then copy the output to the repo root.

## Eligibility feed
The 7 AM eligibility rotation publishes **Eligibility History.csv** (one row per Sohar pull). A Power Automate flow hands it to the tracker: the add-in POSTs `{"tracker":"ESK"}` (as `text/plain`, so there's no CORS preflight) and reads the CSV text from the flow's Response (raw text, or JSON with it in `body`, `csv` or `content`). An admin pastes the flow's trigger URL in **Settings › Eligibility feed**; it's saved as the `EligibilityFeedUrl` row of `tblNotifyConfig`, like the Teams `WebhookUrl`.

The pane pulls once a day when the workbook opens (the day is kept in `tblNotifyConfig › EligibilityPulledOn`), and anyone can select **Pull eligibility** (Settings › Imports, or a patient's page). Rows for MRNs on the Master Patient Tracker (digits, leading zeros ignored) replace the hidden **Eligibility Checks** sheet. For each patient, the newest eligible verification (else the newest pull) fills their `tblBenefits` row with Source "Eligibility feed"; Sent As, Sohar ID, Sohar Link, Related Sohar IDs, Projected Patient Responsibility and Estimate Note are added at the right end of the table if missing. BIDF columns aren't touched. Each patient's page lists every pull, newest first, with links to Sohar. If the feed can't be read, nothing is changed.

## Shared settings (ESK Shared sheet)
ESK Tracker and ESK Bays are separate add-ins with their own Ids, and Office keeps document settings per add-in, so neither could read what the other saved: ESK Bays never got the bay map, the techs' vitals or bays, and the tracker never got a provider's dose sign-off or PRN order. Every shared `esk*` setting (bay map and clinic settings in `eskShared`, visits in `eskVisit`, tasks, notes, appt list fingerprint…) now lives on the hidden **ESK Shared** sheet, one row per key: A key, B stamp (changes on every save), C parts, D on the JSON in 30,000-character pieces. Panes poll the stamps every 2 s and read a row only when its stamp changes; an edit to the sheet is read at once. The first time each add-in opens after this change, its own settings are copied over (a key the sheet already has keeps the sheet's values and gains what this add-in had that it lacked; visits merge field by field). Each add-in still keeps a copy in its own settings. Without the sheet (for example a protected workbook structure) the add-in uses its own settings, as before. Office's own `Office.AutoShowTaskpaneWithDocument` stays per add-in.

## Visit log
Every visit action (room time, dose time, bay, initial / 40-min / final vitals, rechecks, PRN meds, discharge, moves) is written to the hidden **ESK Visit Log** sheet as it's saved, one row per change, with the MRN, the time and who recorded it. It's a backup: if a visit's vitals or times go missing from the tracker (for example two people saving at once), the next refresh fills the blanks back in from the log. It never overwrites what's there. A patient's rows stay until the day of their next appointment, then they're cleared.

## Teams notifications
Vitals cards go to each provider's Teams group chat through one Power Automate flow, **ESK Tracker to Teams**, shared by every office (built from blank, not the Teams Workflows template, so there's no "used a Workflow template" footer). The add-in decides whether to send, builds the card (title, stage, vitals with a bold HIGH / EMERGENCY / LOW flag, sent by, and an @mention of the provider) and names the chat; the flow only posts it as Workflows.

Set it up in **Settings › Teams notifications**: paste the flow's HTTP URL once, then for each provider their group chat link (the chat's Teams web address), the name to @mention and their work email. A workbook set up for the older one-workflow-per-provider setup shows an **Update for the new Teams flow** button that converts its Settings sheet in place. **Send test** posts a sample card to a provider's chat.

Thresholds are constants in `notify.js`: emergency at or above 180/120, HIGH over 140/90, LOW under 92% SpO2.

When the tracker shows **Needs MADRS** for the patient (from 30 days before auth expiry until Last MADRS falls in that window), the card adds a line under the vitals: `MADRS: DUE · auth exp 12/5 · last 10/2`, with DUE in bold red.

### Start administration from Teams
With **Start administration button on cards** turned on (Settings › Teams notifications, off by default), a card for a visit with no dose time yet gets a **Start administration** button. Emergency cards never get it. A tap sets the visit's dose time to the tap time, which starts the 40-min and final vitals timers. It never replaces a dose time already entered, and the tech's time in Administer still replaces it. The pane shows a toast, and the visit log records the dose time as "Start administration in Teams" with the name of whoever tapped. Anyone in the chat can tap, and only the first tap counts.

Turning the setting on adds **tblNotifyActions** (`Ref | Action | By | At | Applied`) at X1 on the Settings sheet. The Excel pane checks it every 15 s while a visit today has a card with the button and no dose time, then fills **Applied** (`Dose time set 10:42`, or why the tap was skipped).

The flow needs these changes **before** you turn the setting on. A plain "Post card" can't take a tap, so the button would show an error.
1. In **Apply to each** over `chats`, add a **Condition**: `triggerBody()?['wait']` is equal to `true`.
2. **If yes:** Microsoft Teams › **Post adaptive card and wait for a response**. Post as Flow bot, in a Group chat, Group chat = `item()`, Message = the card (`triggerBody()?['card']`), Update message `Administration started`, Should update card = Yes. In its Settings, set Timeout to `PT4H`. Then Excel Online (Business) › **Add a row into a table** in that office's workbook, with a Switch on `triggerBody()?['clinic']` and one case per office. Table `tblNotifyActions`. Ref = `body('Post_adaptive_card_and_wait_for_a_response')?['data']?['ref']`, Action = `…?['data']?['action']`, By = `…?['responder']?['displayName']`, At = `…?['responseTime']`. Leave Applied blank.
3. **If no:** the existing **Post card in a chat or channel**, unchanged.
4. Turn on concurrency for the Apply to each, so a waiting card doesn't hold up the next chat.

Check the output names (`data`, `responder`, `responseTime`) against one test run, because Power Automate's action names follow the names in your flow.
