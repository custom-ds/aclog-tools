# Amateur Contact Log Tools
Tools to update and visualize N3FJP's Amateur Contact Logbook.

## Import ADIF
Import ADIF will take an ADIF file of a user's logbook on QRZ.com, and merge
it into the N3FJP Amateur Contact Log

This script will update an N3FJP Amateur Contact Logbook with lat/lon information that.
is downloaded in an ADIF file from QRZ.com. In order to get the ADIF file, you must.
be subscribed to QRZ.com, and download your logbook from QRZ.com and select the ADIF format.

The downloaded ADIF file should be placed in this directory, and named 'import.adi'.
When this script is run, it will read the ADIF file and update the logbook with the lat/lon
into the user defined fldOther7 and fldOther8 fields in the logbook. It also stores the QSL
Manager information that is registered with QRZ in the fldOther6 field.

This script requires that you have a QRZ.com account and have a valid subscription.
You will need to provide the path to your logbook file in the .env file. See the
template.env file for examples.

This script is provided as-is and without warranty. Use at your own risk.

**IMPORTANT:**
This script will update your logbook. BE SURE TO BACKUP YOUR LOGBOOK BEFORE RUNNING THIS SCRIPT.
Also, it is recommended to close out of Amateur Contact Logbook before running this script.

## Lat/Lon Lookup
Update N3FJP Amateur Contact Log with Lat/Lon

This script will update an N3FJP Amateur Contact Logbook with missing lat/lon information.
The script will query QRZ.com for the missing lat/lon information and store it in the.
user defined fldOther7 and fldOther8 fields in the logbook. It also stores the QSL Manager
information that is registered with QRZ in the fldOther6 field.

This script requires that you have a QRZ.com account and have a valid subscription.
You will need to provide your QRZ.com username and password in the .env file, as well as
the path to your logbook file. See the template.env file for examples.

This script is provided as-is and without warranty. Use at your own risk.

**IMPORTANT:**
This script will update your logbook. BE SURE TO BACKUP YOUR LOGBOOK BEFORE RUNNING THIS SCRIPT.
Also, it is recommended to close out of Amateur Contact Logbook before running this script.


## Ham Dashboard
The visualization Dashboard.pbix is a Microsoft Power BI dashboard that visualizes the contents
of an Amateur Contact Log logfile. 

## Installation Instructions
The Python scripts require Python 3 to be installed on your PC. They need to have
a couple of standard libaries installed. From a command prompt, run the pip installer
process:

```
pip install -r requirements.txt
```

The scripts should be ready to run by calling them from Python:
```
python import_adi.py
```
or
```
python latlon_lookup.py
```

### Microsoft Power BI
The Ham Dashboard is a Power BI dashboard, which is a Microsoft visuationation tool. The
Power BI Desktop can be downloaded for free from Microsoft for use on your local PC.

Once the Dashboard.pbix has been opened, click on the Get Data from the toolbar and 
select the Microsoft Access .mdb from the Logbook.
