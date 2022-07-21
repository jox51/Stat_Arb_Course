# Stat_Arb_Course
Some additional files made to automate creating the backtesting file as done in the course and auto upload to google sheets.

Requires PyDrive2 and pygsheets. Will need to create to separate secret json files for both. Follow access instructions on their repo to get required JSON access files for both.

PyDrive2:
https://github.com/iterative/PyDrive2

pygsheets:
https://github.com/nithinmurali/pygsheets

PyDrive2 is required in order to upload an excel file to Google Sheets.
pygsheets required to access uploaded Excel sheet and extact/modify spreadsheet as needed.

Basic flow of operations:


    1 - Use pygsheets to create a new excel file
            client = pygsheets.authorize('pygsheets json file')
            client.sheet.create('name you want for your file') 
            
    2 - Use PyDrive2 to initialize a Drive instance, loop through list of spreadsheets, find the title of 
        sheet you use created, once found, grab this sheet ID and replace with your local spreadsheet
        
                sheet_list = drive.ListFile({'q':"mimeType='application/vnd.google-apps.spreadsheet'"}).GetList()
                for file1 in sheet_list:
                   if file1['title'] == 'name of file from step 1':
                     file_id = file1['id']
                      break
                      
                file1 = drive.CreateFile({'id': file_id})
                file1.SetContentFile('your local file you want uploaded')
                file1.Upload()
  
    
    3 - Use pygsheets ('open_by_key') method to access this spreadsheet via it's ID(file_id). Then you can edit and access file as you wish.
    
    
