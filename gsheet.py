import gspread
import itkdb
from oauth2client.service_account import ServiceAccountCredentials
import json

# set up google service account credentials
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name('secret_key.json')

# get itkpd access codes from itkpd_access.json file 
with open('itkpd_access.json', 'r') as file:
  keys = json.load(file)
db_pass1 = keys['ACCESS_CODE1']
db_pass2 = keys['ACCESS_CODE2']

# setup itkdb access
user = itkdb.core.User(access_code1 = db_pass1, access_code2 = db_pass2)
client = itkdb.Client(user=user)

# Get sheets for each module type
file = gspread.authorize(creds)
workbook = file.open("Vancouver ITk Production Tracker")
R1_sheet = workbook.worksheet('R1')
R2_sheet = workbook.worksheet('R2')
R4_sheet = workbook.worksheet('R4')
R5_sheet = workbook.worksheet('R5')


######## R1 modules 

# reset
new_locations = []
new_stages = []
stage_cells = 'W3:W200'
location_cells = 'X3:X200'

print("UPDATING STAGES AND LOCATIONS OF R1 MODULES")
for cell in R1_sheet.range('B3:B200'):
    print(cell.value)
    mod = cell.value
    if cell.value != "": 
      current_component = client.get("getComponent", json={"component": mod})
      current_mod_location = current_component['currentLocation']['code']
      current_mod_stage = current_component['currentStage']['code']

      new_stages.append([str(current_mod_stage)])
      new_locations.append([str(current_mod_location)])
    else:
      new_stages.append([""])
      new_locations.append([""])
R1_sheet.update(range_name=stage_cells, values=new_stages)
R1_sheet.update(range_name=location_cells, values=new_locations)   


######## R2 modules 

# reset
new_locations = []
new_stages = []
stage_cells = 'W3:W200'
location_cells = 'X3:X200'

print("UPDATING STAGES AND LOCATIONS OF R2 MODULES")
for cell in R2_sheet.range('B3:B200'):
    print(cell.value)
    mod = cell.value
    if cell.value != "": 
      current_component = client.get("getComponent", json={"component": mod})
      current_mod_location = current_component['currentLocation']['code']
      current_mod_stage = current_component['currentStage']['code']

      new_stages.append([str(current_mod_stage)])
      new_locations.append([str(current_mod_location)])
    else:
      new_stages.append([""])
      new_locations.append([""])
R2_sheet.update(range_name=stage_cells, values=new_stages)
R2_sheet.update(range_name=location_cells, values=new_locations)     

######## R4 modules 

# reset
new_locations = []
new_stages = []
stage_cells = 'Y3:Y202'
location_cells = 'Z3:Z202'

# Split module - go through each ringmodule to get the stage 
print("UPDATING STAGES OF R4 RING-MODULES")
for cell in R4_sheet.range('B3:B202'):
    print(cell.value)
    mod = cell.value
    if cell.value != "": 
      current_component = client.get("getComponent", json={"component": mod})
      current_mod_stage = current_component['currentStage']['code']
      new_stages.append([str(current_mod_stage)])
    else: 
      new_stages.append([""])
R4_sheet.update(range_name=stage_cells, values=new_stages)

# Split module - go through each half module to get the location 
print("UPDATING LOCATIONS OF R4 HALF-MODULES")
for cell in R4_sheet.range('D3:D202'):
    print(cell.value)
    mod = cell.value
    if cell.value != "": 
      current_component = client.get("getComponent", json={"component": mod})
      current_mod_location = current_component['currentLocation']['code']
      new_locations.append([str(current_mod_location)])
    else: 
      new_locations.append([""])
R4_sheet.update(range_name=location_cells, values=new_locations)    


######## R5 modules 
new_locations = []
new_stages = []
stage_cells = 'Y3:Y202'
location_cells = 'Z3:Z202'

# Split module - go through each ringmodule to get the stage 
print("UPDATING STAGES OF R5 RING-MODULES")
for cell in R5_sheet.range('B3:B202'):
    print(cell.value)
    mod = cell.value
    if cell.value != "": 
      current_component = client.get("getComponent", json={"component": mod})
      current_mod_stage = current_component['currentStage']['code']
      new_stages.append([str(current_mod_stage)])
    else: 
      new_stages.append([""])
R5_sheet.update(range_name=stage_cells, values=new_stages)         


# Split module - go through each half module to get the location 
print("UPDATING LOCATIONS OF R5 HALF-MODULES")
for cell in R5_sheet.range('D3:D202'):
    print(cell.value)
    mod = cell.value
    if cell.value != "": 
      current_component = client.get("getComponent", json={"component": mod})
      current_mod_location = current_component['currentLocation']['code']
      new_locations.append([str(current_mod_location)])
    else: 
      new_locations.append([""])
R5_sheet.update(range_name=location_cells, values=new_locations)    


# Align everything 
R1_sheet.format("W3:Z202", {"horizontalAlignment": "CENTER"})
R1_sheet.format("W3:Z202", {"verticalAlignment": "MIDDLE"})
R1_sheet.format("X3:Y202", {"horizontalAlignment": "CENTER"})
R1_sheet.format("X3:Y202", {"verticalAlignment": "MIDDLE"}) 

R2_sheet.format("W3:Z202", {"horizontalAlignment": "CENTER"})
R2_sheet.format("W3:Z202", {"verticalAlignment": "MIDDLE"})
R2_sheet.format("X3:Y202", {"horizontalAlignment": "CENTER"})
R2_sheet.format("X3:Y202", {"verticalAlignment": "MIDDLE"}) 

R4_sheet.format("Z3:Z202", {"horizontalAlignment": "CENTER"})
R4_sheet.format("Z3:Z202", {"verticalAlignment": "MIDDLE"})
R4_sheet.format("Y3:Y202", {"horizontalAlignment": "CENTER"})
R4_sheet.format("Y3:Y202", {"verticalAlignment": "MIDDLE"}) 

R5_sheet.format("Z3:Z202", {"horizontalAlignment": "CENTER"})
R5_sheet.format("Z3:Z202", {"verticalAlignment": "MIDDLE"})
R5_sheet.format("Y3:Y202", {"horizontalAlignment": "CENTER"})
R5_sheet.format("Y3:Y202", {"verticalAlignment": "MIDDLE"}) 
