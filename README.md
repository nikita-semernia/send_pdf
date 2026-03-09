\# send\_pdf



PowerPoint VBA + Python helper for exporting slides to PNG, building a PDF, and sending it to a student via Telegram.



\## Quick start



1\. Install Python 3.12+

2\. Install dependencies:

&nbsp;  py -3.12 -m pip install requests img2pdf winotify

3\. Create local files from examples:

&nbsp;  - configs/config.json

&nbsp;  - configs/students.json

4\. Open the PowerPoint source presentation

5\. Run the Ribbon button / VBA macro `ExportAndSendPDF`



\## Local-only files



These files must not be committed:

\- configs/config.json

\- configs/students.json

\- logs/\*

\- generated PDF/PNG files



