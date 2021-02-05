import pyodbc
from pathlib import Path as p

loc_db = p.cwd().joinpath('Access')  # Create a path to directory with documents

# temp = (str(p.joinpath(loc_db,"Church_database.accdb")))
try:
    con_string  = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'\
        r'DBQ=C:\\Users\\Ideal\Desktop\\ChurchDataBase-master\\Access\\Church_database.accdb;'
    conn = pyodbc.connect(con_string)
    print("connected")
    # the above path string is the only way it will connect do not change its styokle speciofocally (\\)

    # Add activiteis for the program to enter 

except pyodbc.Error as e:
    print("there was an error" , e)
    

def enter_data(connection,values):
    cursor = con.cursor()

    #data in tuple 

    cursor.execut(f'INSERT INTO database{}\
                         VALUES{}  )
    con.commit()



    





