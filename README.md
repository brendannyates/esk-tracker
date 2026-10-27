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

## Visit log
Every visit action (room time, dose time, bay, initial / 40-min / final vitals, rechecks, PRN meds, discharge, moves) is written to the hidden **ESK Visit Log** sheet as it's saved, one row per change, with the MRN, the time and who recorded it. It's a backup: if a visit's vitals or times go missing from the tracker (for example two people saving at once), the next refresh fills the blanks back in from the log. It never overwrites what's there. A patient's rows stay until the day of their next appointment, then they're cleared.

## Teams notifications
Vitals cards go to each provider's Teams group chat through one Power Automate flow, **ESK Tracker to Teams**, shared by every office (built from blank, not the Teams Workflows template, so there's no "used a Workflow template" footer). The add-in decides whether to send, builds the card (title, stage, vitals with a bold HIGH / EMERGENCY / LOW flag, sent by, and an @mention of the provider) and names the chat; the flow only posts it as Workflows.

Set it up in **Settings › Teams notifications**: paste the flow's HTTP URL once, then for each provider their group chat link (the chat's Teams web address), the name to @mention and their work email. A workbook set up for the older one-workflow-per-provider setup shows an **Update for the new Teams flow** button that converts its Settings sheet in place. **Send test** posts a sample card to a provider's chat.

Thresholds are constants in `notify.js`: emergency at or above 180/120, HIGH over 140/90, LOW under 92% SpO2.
