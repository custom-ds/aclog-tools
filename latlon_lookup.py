import requests
import xml.etree.ElementTree as ET
import pyodbc

import urllib
from dotenv import load_dotenv
import os


def main():

    print("")
    print("")
    print("Update N3FJP Amateur Contact Log with Lat/Lon")
    print("")    
    print("This script will update an N3FJP Amateur Contact Logbook with missing lat/lon information.")
    print("The script will query QRZ.com for the missing lat/lon information and store it in the.")
    print("user defined fldOther7 and fldOther8 fields in the logbook. It also stores the QSL Manager")
    print("information that is registered with QRZ in the fldOther6 field.")
    print("")
    print("This script requires that you have a QRZ.com account and have a valid subscription.")
    print("You will need to provide your QRZ.com username and password in the .env file, as well as")
    print("the path to your logbook file. See the template.env file for examples.")
    print("")
    print("This script is provided as-is and without warranty. Use at your own risk.")
    print("")
    print("")
    print("IMPORTANT:")
    print("This script will update your logbook. BE SURE TO BACKUP YOUR LOGBOOK BEFORE RUNNING THIS SCRIPT.")
    print("Also, it is recommended to close out of Amateur Contact Logbook before running this script.")
    print("")

    print("Do you have a backup of your logbook? Press 'Y' to continue or any other key to exit.")
    response = input("Enter 'Y' to continue: ")
    if response.upper() != "Y":
        print("Exiting without updating the logbook.")
        return

    sessionID = getSession()
    #print(f"QRZ.com Session ID: {sessionID}")

    # Connect to the database
    #Get the enviro variables from the .env file
    load_dotenv()

    # Access the variables
    logPath = os.getenv('ACLOG_PATHFILENAME')

    # Define the connection string
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        r"DBQ=" + logPath + ";"
    )

    #Get all of the callsigns from rows that other7 or other8 is null
    #SQL statement
    topX = "" #"TOP 10"        # Uncomment to test a small sample size
    sql = f"SELECT {topX} [fldCall] FROM tblContacts WHERE [fldOther7] = '' OR [fldOther8] = '';"

    # Try to connect to the database
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
    except Exception as e:
        print(f"Failed to connect to the database: {e}")
        print("")
        print("Verify that a .env file exists, and that the correct path to the logbook is defined.")
        return

    # Execute the SQL statement
    cursor.execute(sql)

    # Fetch all the rows
    rows = cursor.fetchall()

    #Display how many rows are in the result set
    print(f"There are {len(rows)} contact(s) that were found in the logbook that are missing lat/lon.")

    #Confirm that the user wants to update the logbook
    print("Would you like to update the logbook with the missing lat/lon?")
    response = input("Enter 'Y' to update the logbook: ")
    if response.upper() != "Y":
        print("Exiting without updating the logbook.")
        return

    # Close the cursor and the connection
    cursor.close()

    # Iterate over the rows and update the Other7 and Other8 columns
    for row in rows:
        callsign = row[0]
        
        print(f"Processing callsign: {callsign}")
        result = getLatLon(sessionID, callsign)
        if result:
            latitude, longitude, qslVia = result
            print(f"  The lat/lon found were: {latitude}, {longitude}. Updating the logbook.")

            #Clean up the qslVia to avoid sql injection
            qslVia = qslVia.replace("'", "").replace('"', "").replace(";", "").replace("--", "").replace("/*", "").replace("*/", "")

            # Connect to the database
            cursor = conn.cursor()

            # Update the Other7 and Other8 columns
            sql = f"UPDATE tblContacts "
            sql = sql + f"SET [fldOther7] = '{latitude}', [fldOther8] = '{longitude}', [fldOther6] = '{qslVia}' "
            sql = sql + f"WHERE [fldCall] = '{callsign}' AND [fldOther7] = '' AND [fldOther8] = '';"
            #print(sql)
            cursor.execute(sql)

            # Commit the transaction
            conn.commit()

            # Close the cursor and the connection
            cursor.close()
        else:
            print(f"Failed to retrieve lat/lon for callsign {callsign}")
        
        print("")
        print("")
        
    #Close out the database connection
    conn.close()

def getSession():

    # Make a call to the QRZ API to get a session key
    # Use your QRZ username and password here

    # Load environment variables from .env file
    load_dotenv()

    # Access the variables
    username = os.getenv('QRZ_USERNAME')
    password = os.getenv('QRZ_PASSWORD')


    #urlencode the password
    password = urllib.parse.quote(password)
    
    response = requests.get(f"https://xmldata.qrz.com/xml/current/?username={username};password={password};agent=python3.9")

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the XML response to extract the session key
        # You may need to use an XML parsing library like xml.etree.ElementTree for this
        # Extract the session key from the parsed XML
        
        #get the body of the response
        xml = response.text
        #print(xml)
        # Parse the XML data
        root = ET.fromstring(xml)

        # Define the namespace
        namespace = {'ns': 'http://xmldata.qrz.com'}

        # Find the Key element and get its value
        sessionID = root.find('.//ns:Key', namespace).text

        # Return the session key
        return sessionID
    else:
        # Handle the case when the request fails
        print("Failed to retrieve session key from QRZ API.")
        print("Verify that your QRZ username and password are correct in the .env parameters file.")
        print("See the template.env file for examples.")
        return None
def getLatLon(sessionID, callsign):
    
    response = requests.get(f"https://xmldata.qrz.com/xml/current/?s={sessionID};callsign={callsign};agent=python3.9")

    # Check if the request was successful
    if response.status_code == 200:
        #get the body of the response
        xml = response.text
        #print(xml)
        # Parse the XML data
        root = ET.fromstring(xml)

        # Define the namespace
        namespace = {'ns': 'http://xmldata.qrz.com'}

        errorMsg = root.find('.//ns:Error', namespace)
        if errorMsg is not None:
            print(f"Error: {errorMsg.text}")
            return None
        
        # Find the Key element and get its value
        latitude = root.find('.//ns:lat', namespace).text
        longitude = root.find('.//ns:lon', namespace).text

        #check to see if there is a qslmgr element
        qslVia = ""
        if root.find('.//ns:qslmgr', namespace) is not None:
            qslVia = root.find('.//ns:qslmgr', namespace).text

        #return the latitude and longitude as a tuple
        return (latitude, longitude, qslVia)

    else:
        # Handle the case when the request fails
        print("Failed to retrieve Lat/Lon from QRZ.")
        return None

if __name__ == "__main__":
    main()