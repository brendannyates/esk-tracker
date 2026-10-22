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

## Payer credential rules
In Settings, select **Import payer rules…** and choose the shared **Payer Credential Rules** file (CSV or xlsx, in the OneDrive Add-In Manifests folder). It's written to an editable **Payer Credential Rules** sheet in the workbook; change a rule there and the panel picks it up on refresh. The list itself is never committed here (this repo is public).

For Active / Pending Start patients whose payer (the tracker's Insurance, else the Sohar payer) is on the list, each booked visit is checked against its rendering provider's credential (Templates & chairs, else the name). When the payer doesn't let that credential bill the visit (start, first treatment or observation), the visit gets a red chip, Today shows a Billing alert, and Admin to-do lists it. Besse patients whose payer isn't Besse-approved get a warning chip; a blank Besse cell shows "Besse approval pending". These are flags only: ordering and kit assignment don't change. When a payer has rules for more than one state, Settings › Clinic › State picks one. Payers not on the list, or no sheet, show nothing. `payer-rules.js` is shared with the Clinic Template Tool and Sohar Eligibility; change all three together.

## Visit log
Every visit action (room time, dose time, bay, initial / 40-min / final vitals, rechecks, PRN meds, discharge, moves) is written to the hidden **ESK Visit Log** sheet as it's saved, one row per change, with the MRN, the time and who recorded it. It's a backup: if a visit's vitals or times go missing from the tracker (for example two people saving at once), the next refresh fills the blanks back in from the log. It never overwrites what's there. A patient's rows stay until the day of their next appointment, then they're cleared.
