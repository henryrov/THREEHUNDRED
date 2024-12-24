What?
=====
A script that takes in a Workday Student saved course schedule .xlsx file and makes an iCalendar file based on it. Specifically meant for UBC courses (assumed formatting and timezone).

How?
====
Navigate to the saved schedule you want to use and click the "Export to Excel" button to the top right of the table. Run the script on the downloaded file:
```
python3 p01.py <filename>
```
Verify the detected course schedule. If it looks right, import the generated "out.ics" file into your calendar client.