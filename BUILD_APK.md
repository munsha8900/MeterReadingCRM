# Getting Home Solar Monitoring onto your Android phone

You have two ways. Option A works in 2 minutes with no build at all.
Option B produces a real installable .apk file.

------------------------------------------------------------------
## Option A — Install the app instantly (recommended)
------------------------------------------------------------------
The app is a single offline web app. No build, no Play Store.

1. Copy `MeterCRM.html` to your phone (email it to yourself, put it on
   Google Drive, or transfer by USB).
2. Open it in Chrome on the phone.
3. Tap the ⋮ menu (top-right) → "Add to Home screen".
4. It now has its own icon and opens full-screen, like an app. It works
   offline and saves your edits on the phone.

That's it. This is the fastest path and it already has all your data.

------------------------------------------------------------------
## Option B — Build a real .apk (no Android Studio needed)
------------------------------------------------------------------
This uses GitHub's free cloud builders to compile an installable APK.
The Android build can't run on the machine that generated this package,
so it runs in the cloud instead. Steps:

1. Create a free GitHub account if you don't have one.
2. Make a new repository and upload the WHOLE contents of this folder
   to it (keep the folders `android/` and `.github/` exactly where they
   are — they must sit at the top level of the repo).
3. On the repo page open the "Actions" tab. If prompted, enable Actions.
4. Pick "Build Android APK" on the left → "Run workflow" → Run.
5. Wait ~5 minutes. When it finishes (green tick), open the run and
   download the artifact named "MeterCRM-apk". Inside is `app-debug.apk`.
6. Copy that .apk to your phone and tap it to install. You'll be asked to
   allow "install from unknown sources" the first time — allow it.

To change the app later, edit `MeterCRM.html`, copy it over
`android/app/src/main/assets/index.html`, push, and run the workflow again.

Notes:
- The APK is a debug build (fine for personal use / sideloading). For the
  Play Store you'd sign a release build — ask if you need that.
- Package id: com.markfold.metercrm · app name: Home Solar Monitoring

------------------------------------------------------------------
## Option C — Turn it into an APK online
------------------------------------------------------------------
If you'd rather not use GitHub: host `MeterCRM.html` anywhere with https
(e.g. a free Netlify "drop"), then paste that URL into
https://www.pwabuilder.com — it packages an Android APK for you.
