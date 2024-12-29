import pyodbc
from datetime import datetime

from dotenv import load_dotenv
import os

def main():


    print("")
    print("")
    print("Import ADIF file into N3FJP Amateur Contact")
    print("")
    print("This script will update an N3FJP Amateur Contact Logbook with lat/lon information that.")
    print("is downloaded in an ADIF file from QRZ.com. In order to get the ADIF file, you must.")
    print("be subscribed to QRZ.com, and download your logbook from QRZ.com and select the ADIF format.")
    print("")
    print("The downloaded ADIF file should be placed in this directory, and named 'import.adi'.")
    print("When this script is run, it will read the ADIF file and update the logbook with the lat/lon")
    print("into the user defined fldOther7 and fldOther8 fields in the logbook. It also stores the QSL")
    print("Manager information that is registered with QRZ in the fldOther6 field.")
    print("")
    print("This script requires that you have a QRZ.com account and have a valid subscription.")
    print("You will need to provide the path to your logbook file in the .env file. See the")
    print("template.env file for examples.")
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

    #Get the enviro variables from the .env file
    load_dotenv()

    # Access the variables
    logPath = os.getenv('ACLOG_PATHFILENAME')

    # Define the connection string
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        "DBQ=" + logPath + ";"
    )

    # Connect to the database
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Read the file content from log.txt
    with open("import.adi", "r") as file:
        file_content = file.read()

    # Extract fields from the file content
    rows = extract_fields(file_content)

    # Print the results
    for row in rows:
        sql = updateMDB(row)
        print(row)
        print(sql)
        # Execute the update statement
        cursor.execute(sql)
        conn.commit()

    # Close the connection
    cursor.close()
    conn.close()



def updateMDB(row):

    # Define the update statement
    update_statement = "UPDATE tblContacts SET "
    update_statement += f"fldOther6 = '{row['qsl_via']}', "
    update_statement += f"fldOther7 = '{row['lat']}', "
    update_statement += f"fldOther8 = '{row['lon']}' "
    update_statement += f"WHERE fldCall = '{row['call']}' AND fldOther7 = '' AND fldOther8 = '' "

    #Format the qso_date in mm/dd/yyyy format
    tmpDate = row['qso_date'].strftime("%Y/%m/%d")
    update_statement += f"AND fldDateStr = '{tmpDate}';"

    return update_statement


def extract_fields(file_content):
    records = file_content.split("<eor>")
    extracted_data = []

    for record in records:
        if record.strip():
            call = lat = lon = qso_date = qsl_via = None

            for line in record.split("\n"):
                if "<call:" in line:
                    call = line.split(">")[1]
                elif "<lat:" in line:
                    lat = format_decimal(convert_to_decimal(line.split(">")[1]))
                elif "<lon:" in line:
                    lon = format_decimal(convert_to_decimal(line.split(">")[1]))
                elif "<qso_date:" in line:
                    qso_date = convert_to_date(line.split(">")[1])
                elif "<qsl_via:" in line:
                    qsl_via = line.split(">")[1]

            if qsl_via is None:
                qsl_via = ""
            else:
                #strip out any quotes or double quotes from the qsl_via field
                qsl_via = qsl_via.replace("'", "")
                qsl_via = qsl_via.replace('"', "")

            extracted_data.append({
                "call": call,
                "lat": lat,
                "lon": lon,
                "qso_date": qso_date,
                "qsl_via": qsl_via
            })

    return extracted_data

def convert_to_date(date_str):
    return datetime.strptime(date_str, "%Y%m%d").date()

def convert_to_decimal(coord):
    direction = coord[0]
    degrees, minutes = map(float, coord[1:].split())
    decimal = degrees + minutes / 60
    if direction in 'SW':
        decimal *= -1
    return decimal

def format_decimal(decimal):
    return f"{decimal:.5f}"



if __name__ == "__main__":
    main()