HDASCO EXTRACTOR - PYTHON + GOOGLE CHROME
=========================================

IMPORTANT
---------
This is a starter build for live testing against your HDASCO account. The browser launcher,
Chrome connection, desktop UI, Excel formatting, and flexible XML parser are included.
HDASCO's Anthem callbacks and XML structure can vary; the GET XML and page-pagination
automation must be validated on the actual logged-in page before relying on bulk exports.
The tool does not bypass login or access controls.

REQUIREMENTS
------------
- Windows 11
- Google Chrome installed
- Python 3.11 or 3.12 recommended (64-bit)
- Internet access to install Python packages
- An authorized HDASCO account

FIRST RUN
---------
1. Extract this folder to a normal writable folder, for example:
   C:\HDASCO_Python_Extractor
2. Install Python from python.org. During setup, tick "Add Python to PATH".
3. Double-click run_app.bat.
4. In the app, click "Open HDASCO in dedicated Chrome".
5. Sign in to HDASCO in the new Chrome window.
6. Navigate to BR List or BL List and confirm the list is visible.
7. Return to the app and click "Connect to Chrome".
8. Run one extraction at a time. Start with Pages = 1 for testing.

OUTPUT
------
Documents\HDASCO_Exports
  BR_List.xlsx
  XML\*.xml
  HDASCO_BL_Export.xlsx
  Container_History.xlsx

BUILD EXE
---------
If Python is already installed, double-click build_exe.bat. The executable will be created at:
  dist\HDASCO_Extractor.exe

If you do NOT have Python, follow BUILD_EXE_WITHOUT_PYTHON.txt to build the EXE on a cloud
Windows runner using GitHub Actions. That method still requires a GitHub account and uploading
the project files, but does not require installing Python on your own PC.

NOTES
-----
- The app uses a separate Chrome profile:
  %LOCALAPPDATA%\HDASCOExtractor\ChromeProfile
- This is intentionally not your normal Chrome profile.
- Chrome must remain open while extracting.
- Do not close the dedicated Chrome window while the app is running.
- If the site expires the session, log in again and retry.
- The current site may use callback navigation that requires adjustment. If a run reports
  that XML was not captured, send the exact message and describe what Chrome did.
- Pricing and package extraction depend on the actual tag names in the downloaded XML.
  If a worksheet is empty, inspect a sample XML and update the parser to match its tags.


UPDATED EXCEL OUTPUT: The exporter now creates one worksheet named "All BL Data" in HDASCO_BL_Export.xlsx. The RecordType column labels rows as BL Summary, Container, Package, or Pricing. Use Excel filters to view each record type.
