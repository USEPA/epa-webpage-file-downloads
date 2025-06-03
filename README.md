This repository contains a script which can be used to download documents from EPA webpages. The user will provide the URL for an EPA webpage or add a list of URLs to the urls.txt file. The script will then download all xlsx, xls, pdf, docx, doc, zip, csv files from the page into a /downloads folder in the same folder as the script.

Users can clone the repo or download it as a zip file and extract it. Once downloaded, double click the "Double Click Here.bat" file to run the script.

If there is nothing in the urls.txt file, the script will prompt you to enter a URL to download files:

<img src="/img/EnterURL.png" alt="Screenshot of the prompt asking users to enter a URL" width="250">

Click "OK" and the script will check for files on the page. It will then prompt you with "You are about to download files from: _USER-ENTERED EPA URL_. Do you want to continue?"

<img src="/img/WantToContinue.png" alt="Screenshot of the prompt asking users if they want to continue downloading files" width="250">

Click "Yes" and the script will begin downloading files to a /downloads folder in the same location as the script file.
