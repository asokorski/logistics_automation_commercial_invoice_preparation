import cx_Oracle
cx_Oracle.init_oracle_client(lib_dir=r"C:\Oracle\instantclient_21_11")
import os
import shutil
from database_credentials import source_directory, dsn_tns, username, password

print("")
print("The program is designed to copy and rename the commercial invoices for the export orders")
print("v 1.0.0 Developed by Adrian Sokorski to make your life easier")
print("")
print("Current program configuration searches for the following countries:")
print("NORWAY, SWITZERLAND, ANDORRA, MALTA, MARTINIQUE, SAINT BARTHELEMY, SAINT-MARTIN, SAINT PIERRE AND MIQUELON, MAYOTTE, GUYANA, JERSEY, ICELAND, CYPRUS, GUERNSEY, GUADELOUPE, RÉUNION, CANARY ISLANDS (visible as SPAIN)")
print("If you wish to add more countries, kindly ask Adrian")
print("")

script_directory = os.path.dirname(os.path.realpath(__file__))
destination_directory = None

while True:

    connection = None
    files_processed = False

    try:
        connection = cx_Oracle.connect(user=username, password=password, dsn=dsn_tns)

        search_string = input("Enter the despatch date in format DD-MM-YYYY or type 'exit' if you want to close the program: ")

        if search_string.lower() == 'exit':
            break

        query = """
        SELECT DISTINCT
        HLGESOP.gsrodp as Order_number,
        HLGESOP.gsrcol as TrackingID,
        HLPISOP.pylpie as Country
        FROM
        HLGESOP
        inner join HLPRENP on gsrodp = perodp
        left join hlodpep on hlodpep.oerodp=hlprenp.perodp
        left join hladgpp on adgidl1adr = (oecact||oecdpo||LPAD(oenann, 2, 0)||LPAD(oenodp,9,0))
        left join HLPISOP on hlpisop.pycnpi = hladgpp.adgcnpi
        where HLPRENP.pecrgc in ('UPS EXPRES', 'UPSSTDEXP', 'UPS', 'UPSEXP')
        and HLPISOP.pylpie in ('NORWAY', 'SWITZERLAND', 'ANDORRA', 'MALTA', 'MARTINIQUE', 'SAINT BARTHELEMY', 'SAINT-MARTIN', 'SAINT PIERRE AND MIQUELON', 'MAYOTTE', 'GUYANA', 'JERSEY', 'ICELAND', 'CYPRUS', 'GUERNSEY', 'GUADELOUPE', 'RÉUNION', 'SPAIN') --only Canary Islands are Spain and we use UPS for them
        and LPAD(GSJCRE, 2, 0) || '-' || LPAD(GSMCRE, 2, 0) || '-' || GSSCRE || GSACRE = :1
        """

        print('Checking the data in the database...')

        cursor = connection.cursor()
        cursor.execute(query, (search_string,))
        results = cursor.fetchall()
 
        num_rows = len(results)
        print(f"Number of export orders found for that date: {num_rows}")

        if len(results) == 0:
            print("No orders for the given date")
            exit()
            
        # Display the results
        print("ORDER_NUMBER\tTRACKINGID\tCountry")
        for row in results:
            print("\t".join(map(str, row)))

        while True:
            copy_files = input("Do you want to copy and rename the files? Type 'yes' or 'no': ").strip().lower()

            if copy_files == "yes":
                # destination directory here
                destination_directory = os.path.join(script_directory, search_string)

                if not os.path.exists(destination_directory):
                    os.makedirs(destination_directory)
                    print(f"Created folder: {search_string} in: {script_directory}")
                    print('The files are being copied to the folder, please wait...')

                # copy and rename here
                files_processed = 0  #successfully processed

                for order_number, tracking_id, country in results:
                    for filename in os.listdir(source_directory):
                        if order_number in filename:
                            source_path = os.path.join(source_directory, filename)
                            new_filename = f"{tracking_id}_{filename}"
                            destination_path = os.path.join(destination_directory, new_filename)
                            shutil.copy2(source_path, destination_path)
                            files_processed += 1

                if files_processed == len(results):
                    print("All results copied and renamed")
                else:
                    print("Part of the results (or no results) copied and renamed")
                break

            elif copy_files == "no":
                print("Files not copied and renamed.")
                break 

            else:
                print("Invalid input. Please enter ' yes' or 'no'.")

    except Exception as e:
        print("Error:", e)

    finally:
        if connection:
            connection.close()
