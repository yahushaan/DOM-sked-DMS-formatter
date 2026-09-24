# MDV Schedule Converter — Internal Web Application

This application converts the iFlight Lite Daily OPS Schedule PDF into the required MDV Schedule Excel format.

## Tested logic

The supplied sample PDF is a 3-page iFlight Lite report with aircraft sections such as `AT46 Regn. 8QIAV` and rows containing Flt nbr, From, STD, BlkOff, BlkIn, STA and To.

The application:

- Reads the operating date from the schedule header (not the printed timestamp).
- Tracks aircraft type and registration by section.
- Reads rows even when the first row contains a leading date/registration before the flight number.
- Handles blank Block Off / Block In columns, which collapse in PDF text extraction.
- Converts `Q2706` → `Q2 706`, `Q2220` → `Q2 220`, etc.
- Groups consecutive sectors into one aircraft rotation until the aircraft reaches MLE.
- Combines paired flight numbers using the required style: `Q2 300/1`, `Q2 2232/3`, `Q2 2448/9`.
- Builds routing such as `MLE DRV IFU MLE`.
- Uses STD when the rotation starts at MLE and STA when it ends at MLE.
- Uses BlkOff/BlkIn as ATD/ETA when those fields are present in the PDF.
- Sorts arrivals first by STA, then departures by STD.
- Excludes A320/international rotations by default because the supplied required Schedule template excludes those three A320 rotations. A checkbox can include them if needed.
- Preserves the provided Excel template formatting.

## Run on Windows

1. Install Python 3.11 or newer.
2. Open Command Prompt in this folder.
3. Run:

   `python -m pip install -r requirements.txt`

4. Double-click `run.bat`, or run:

   `python app.py`

5. Open `http://localhost:8080`.
6. Upload the daily PDF and click **Convert to Excel**.

## Network/internal use

The server listens on `0.0.0.0:8080`, so it can be made available to other computers on the same internal network using the host PC's internal IP address, subject to the organization's firewall rules.

## Important

The application stores uploaded PDFs and generated Excel files in local `uploads` and `output` folders. Keep the application on a controlled internal machine/network and remove old files according to your organization's retention requirements.
