import os
import pandas as pd

def read_xl(path):
    if not os.path.isfile(path):
        # Is the file there?
        raise FileNotFoundError('Failed to find file ' + path)
    try:
        # Can we read it?
        return pd.read_excel(path).astype('str')
    except pd.parser.CParserError as errorm:
        raise FileNotFoundError('Failed to read file ' + path + ' -e: ' + errorm)


if __name__ == "__main__":
    try:
        # Read Data
        files = os.listdir('data/') # List all files in root directory
        updates = []
        databasen = ''
        for file in files:
            if '.xlsx' in file: # Find all Excel Files
                if 'database' in file.lower():
                    databasen = file # The last Excel file containign database is the main database, shouldonly be 1
                else:
                    updates.append(file) # The rest are files to consider to update
        
        print('The DataBase found is: ' + databasen + '.') # Make sure the database file is the inteded, else quit
        chk = input('Is this correct? (y/n): ')
        if chk.lower() != 'y':
            raise Exception('Make sure the database you want is the only .xlsx file with "database" in its name.')
        database = read_xl('data/' + databasen)

        # Update database with each otehr file
        for upfile in updates:
            # Make sure they itnended this to be a file to update DB with
            chk = input('Do you want to update the DataBase with ' + upfile + '? (y/n): ')
            if chk.lower() == 'y':
                try:
                    update = read_xl('data/' + upfile)

                    # Attach new columns from update to database
                    for field in list(update):
                        if field not in list(database):
                            database[field] = ''

                    # Start Search
                    for indexr, record in update.iterrows():
                        # For each record to be updated
                        if record[0] in database.ix[:, 0].unique(): # Does the record exist?
                            # If so, update
                            index_to_update = database.ix[:, 0][database.ix[:, 0] == record[0]].index[0]
                            for field in list(update):
                                # Update each field independently to avoid overrides on emptys
                                if str(update[field][indexr]) != 'nan':
                                    database.set_value(index_to_update, field, update[field][indexr])
                        else:
                            # Else, create
                            database = database.append(record, ignore_index=True)
                    print('Updated DataBase with ' + upfile)
                except Exception as e:
                    print('Failed to update DataBase with ' + upfile)
                    print('Error: ' + e)
        print('Finished updating!')
        # Save Updated Dataset
        print('Saving Database ...')
        database = database.replace({'nan': ''})
        new_file = False
        filei = 0
        while not new_file: # Make sure the new DB is being saved to a new file to prevent errors.
            if os.path.isfile('data/updated_'+str(filei)+'.xlsx'):
                filei += 1
            else:
                new_file = True
        new_name = 'data/updated_'+str(filei)+'.xlsx'
        database.to_excel(new_name, index=False)
        input('New DataBase saved to ' + new_name)
    except Exception as e:
        print(e)
