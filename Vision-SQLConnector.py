import glob, os, pandas as pd, pyodbc, time

# Keep this file running continuously in the background. Sign-in required at startup.


# Get most recent file
directory = sorted(glob.glob(r'C:\Users\BRADLEYCONRAD\Downloads\Vision Documents\*.txt'), key=os.path.getmtime, reverse=True)[:20]
filenames = [f for f in directory if not f.endswith('Dimensions.txt')]

for filename in filenames:
# Load data
    with open(filename, "r", encoding="utf-16-le") as f:
        lines = f.readlines()[:18]
    data = {( line.split(':')[0].strip() ) : ( (line.split(':', 1)[1]).split('(')[0].strip() )
        for line in lines if ":" in line}

    series = pd.Series(data)

# Prepare data for entry
    series['Part Number'] = (series['File']).split('_')[0]
    series['Batch Number'] = max((series['Order number'].replace(series['Part Number'], '')).split('-'), key=len)
    series['Date and Time'] = series['Date'] + ' ' + series['Time']
    series['Vision Machine Number'] = 'V2' # Change this line for each machine.

# Prepare entry for insertion
    entry = series.loc[[
        'Batch Number',
        'Batch Number',
        'Part Number',
        'Date and Time',
        'Inspection duration',
        'Vision Machine Number',
        'Checked',
        'OK (category 1)',
        'Defect'
        ]]

    print(entry)

# Connect to SQL database
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=lagrangesql.database.windows.net;'
        'DATABASE=lagrangesql;'
        'Authentication=ActiveDirectoryInteractive;'
        'UID=bradley.conrad@fnst.com;'
        )

    cursor = conn.cursor()

# Insert entry into SQL database
    insert_query = '''
        IF NOT EXISTS(
            SELECT * 
            FROM dbo.[Vision Sorting Data]
            WHERE [Batch Number] LIKE ?
        )
        BEGIN
            INSERT INTO dbo.[Vision Sorting Data](
                [Batch Number],
                [Part Number],
                [Date and Time],
                [Inspection Duration], 
                [Vision Machine Number],
                [Yield + Scrap],
                [Yield],
                [Scrap]
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        END
        '''

# Close SQL connection
    cursor.execute(insert_query, *entry)
    conn.commit()
    cursor.close()
    conn.close()

# Wait 5 minutes before re-running program.
    print('\n'
        '\n'
        'Keep this program running in the background.\n'
        'To start program, paste the line below at the bottom of this window:\n'
	'\n'
	r'exec(open(r"C:\Users\BRADLEYCONRAD\Downloads\Vision-SQLConnector.py").read())', '\n'
	'\n'
	'\n'
        'Press Ctrl + C to terminate the program.')
    time.sleep(0.1)
