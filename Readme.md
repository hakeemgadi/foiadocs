# Automated download and OCR of FOIA files

There is a huge number of documents in the [FOIA website](https://foia.state.gov/Search/Search.aspx) (Freedom of Information Act). To make the site more useful for researchers, I have created this piece of code to make it possible to download selectively using search terms and a date range. The code first downloads the files and then OCR the scanned PDF pages, and extract any overlaid text.

# Chrome driver version

Make sure that you have the chromedriver.exe that matches your google chrome version. To find out what version is your Chrome browser, go to the "more vertical" menu in the top left corner of your Chrome browser &rarr Help &rarr About Google Chrome. After finding out your Chrome version, download the matching executable from [ChromeDriver downloads page](https://chromedriver.chromium.org/downloads).