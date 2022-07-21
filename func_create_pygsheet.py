from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import pygsheets

# Opens local host for automatic authentication. PyDrive requires separate secret json
gauth = GoogleAuth()
#gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

# Step 1 - Create pygsheets instance and authorize access
def create_gsheet_file(new_file_name):
  client = pygsheets.authorize('pygsheets_secret.json')
  client.sheet.create('backtest_calculation_pairs5') 
 
  # Step 2 -Instantiates Drive to upload local file. Uses title above to retrieve ID
  # Retrieves spreadsheet files on drive to use in for loop
  # Breaks when file id of matching title file is found
  sheet_list = drive.ListFile({'q':"mimeType='application/vnd.google-apps.spreadsheet'"}).GetList()
  for file1 in sheet_list:
    if file1['title'] == 'backtest_calculation_pairs5':
      file_id = file1['id']
      break
      
  # Step 3 - Replaces file above with the local file
  file1 = drive.CreateFile({'id': file_id})
  file1.SetContentFile(new_file_name)
  print(f'New File Name: {new_file_name}')
  file1.Upload()
  
 
  # Step 4 - Uses pygsheets to retrieve value of replaced spreadsheet
  # Uses index notation to reference both sheets
  sh = client.open_by_key(f'{file_id}')
  n_steps = sh[0] 
  m_revert = sh[1]
  
  # Step 6 - Now that we have access to sheets, retrieve data as needed
  net_profit = n_steps.cell('AF11').value
  avg_wr = n_steps.cell('AF17').value
  net_profit_two = m_revert.cell('AI10').value
  avg_wr_two = m_revert.cell('AI16').value
  print(f'n_steps net_profit: ${net_profit}')
  print(f'n_steps Avg Win Rate: {avg_wr}')
  
  print(f'mean_revert net_profit: ${net_profit_two}')
  print(f'mean_revert Avg Win Rate: {avg_wr_two}')
  
  return net_profit, net_profit_two, avg_wr, avg_wr_two
 
  
  

   
 


#### OLD CODE BELOW__ DISREGARD ####
# created_spreadsheet = pygsheets.Worksheet()
# print(created_spreadsheet)


# titles = client.spreadsheet_titles()
# print(titles)