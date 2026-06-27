========================================================================
  HOME SOLAR MONITORING  -  Home Solar Meter Readings
  by MarkFold Digital
========================================================================
Three ways to run it. All contain the same data from Meter_Readings.xlsx
(14 months, 410 daily readings, plus the DATA HSS quarterly summary).

  MeterCRM.html   THE PHONE / TABLET APP  ←  start here for Android
                  Open in any browser; works offline. On a phone, use
                  "Add to Home screen" to install it like an app.
                  See BUILD_APK.md to make an actual .apk file.

  desktop/        THE WINDOWS DESKTOP APP (.exe)
                  Double-click desktop/build_exe.bat to create
                  desktop/dist/MeterCRM.exe  (needs Python 3 installed).
                  Or double-click desktop/run_app.bat to run it directly.

  android/ + .github/   APK BUILD KIT (used by BUILD_APK.md, Option B)

------------------------------------------------------------------------
WHAT THE APP DOES
------------------------------------------------------------------------
  Dashboard   Units per channel for a month, the actual saving, and a
              chart across every month.
  Readings    Tap any day to edit its meter readings; daily units are
              worked out for you. Add or delete days.
  Quarterly   Each quarter's three months rolled up, with the settlement:
                 (Export - Off-peak) x Buyback  -  Peak x Peak-rate
              plus the actual saving from your summary.
  Setup       Channel names, Peak / Off-peak / Buyback rates, meter ref,
              add/delete months, export CSV, reset to original data.

Your edits are saved on the device. "Reset to original data" in Setup
restores the readings imported from your spreadsheet.
========================================================================
