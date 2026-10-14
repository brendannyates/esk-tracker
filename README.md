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
