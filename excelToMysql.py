import MySQLdb
import xlrd
import numbers
import datetime

def do_the_thing():
    print("STARTING \n")

    print("ESTABLISHING MYSQL CONNECTION \n")

    # Establish a MySQL connection
    database = MySQLdb.connect(
        host="", user="", passwd="", db="")

    print("MYSQL CONNECTION ESTABLISHED \n")

    # Get the cursor, which is used to traverse the database, line by line
    cursor = database.cursor()

    # Create the INSERT INTO sql query
    query = "INSERT INTO Rep (Period, Rep_ID, Region, Sub_Region, Segment, RSE) " \
            "VALUES (%s, %s, %s, %s, %s, %s)"

    print("GETTING EXCEL SHEET \n")

    # Open the workbook and define the worksheet
    book = xlrd.open_workbook("Rep Data.xls")
    sheet = book.sheet_by_name("Rep Data")

    print("PROCESSING EXCEL SHEET \n")

    # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
    for r in range(1, sheet.nrows):

        print("PROCESSING ROW " + str(r) + " \n")
    #datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(r,0).value, book.datemode)) allows you to convert a column into datetime
        period = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(r,0).value, book.datemode))
        rep_id = sheet.cell(r, 1).value
        region = sheet.cell(r, 2).value
        sub_region = sheet.cell(r, 3).value
        segment = sheet.cell(r, 4).value
        rse = sheet.cell(r, 5).value
        
        # Assign values from each row
        values = (period, rep_id, region, sub_region, segment, rse)

        print("ROW PROCESSED. INSERTING INTO DB \n")

        # Execute sql Query
        cursor.execute(query, values)

        print("INSERT COMPLETE \n")

    print("FINALIZING DB TRANSACTION")

    # Commit the transaction
    database.commit()

    print("CLOSING DATABASE CONNECTION \n")
    # Close the cursor
    cursor.close()

    # Close the database connection
    database.close()

    print("DATABASE CONNECTION CLOSED \n")

    # Print results
    print("IMPORT COMPLETE \n")
    columns = str(sheet.ncols)
    rows = str(sheet.nrows)
    print(columns + " COLUMNS AND " + rows + " ROWS PROCESSED")


def main():
    do_the_thing()


if __name__ == "__main__":
    main()
else:
    print(__name__)
