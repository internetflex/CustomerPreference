# Customer Preference Centre
<h3>Define user marketing info schedule and report on dates when it's due</h3>

<b>Usage:-</b><br>
Edit the sample Excel File 'CustomerData.xlxs' with the customer preferences

Note that more frequent notification dates supercede less frequent.
Ie if day of month is selected then this is superceded by weekday selections
The never column is quick way to turn of notifications without losing 
existing pattern of notifications.

Run the program <b>CustomerPreferenceApp.exe startdate [CsvFile]</b>

<b>Program Parameters</b><br>
StartDate - has format DD-MMM-YYYY <br>
Optional Parameters:- CsvFile - output file name created to contain the results<br>

If no csvFile specified then results are output to console. <br>
By design a fixed number of 90 days is displayed or saved from
the startdate specified.
