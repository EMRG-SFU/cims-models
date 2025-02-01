
#!/usr/bin/env python
# coding: utf-8

# ### Set working folder to "cims-models"

# In[ ]:


import os
from os import path
from datetime import datetime
from IPython import get_ipython
import pandas as pd
import numpy as np

import CIMS

# Check current folder
print("The current folder is",os.getcwd())


# In[ ]:


# Change line below to set current folder to "cims-models" 
# Use ".." to move up one folder level
# Use child folder "name" to move into a directory within a parent folder
import os

get_ipython().run_line_magic('cd', '-q ".."')
print("The current folder is now", os.getcwd())




# +++  Definition of model/scenario/
# ### Adjust scenario parameters and optional update files below

# In[ ]:


# Only uncommented regions below will be run in the simulation
region_list = [
    'CIMS',
    'CAN',
    'BC',
    'AB',
    'SK',
    'MB',
    'ON',
    'QC',
    'AT',
]

# Only uncommented years below will be run in the simulation
year_list = [
    ### Historical ###
    2000,
    2005,
    2010,
    2015,
    2020,
    ### Forecast ###
    2025,
    2030,
    2035,
    2040,
    2045,
    2050,
]

# Only uncommented sectors below will be run in the simulation
# Uncomment individual sectors to calibrate one at a time
# Must use Exogenous prices (and optional exogenous demand) file below when calibrating
sector_list = [
    'Coal Mining',
    'Natural Gas Production',
    'Petroleum Crude',
    'Petroleum Refining',
    'Electricity',
    'Biodiesel',
    'Ethanol',
    'Hydrogen',
    'Mining',
    'Industrial Minerals',
    'Iron and Steel',
    'Metal Smelting',
    'Chemical Products',
    'Pulp and Paper',
    'Light Industrial',
    'Construction',
    'Residential',
    'Commercial',
    'Transportation Personal',
    'Transportation Freight',
    'Waste',
    'Agriculture',
    'Forestry',
]



### Base model and standard files below are required for the model to run
model_path = 'csv/model'
# Model files should be located at: cims-models/path/file/file_XX.csv

### Base model to start initialisation
base_model = 'CIMS_base'

### Default values file
default_path = 'csv/defaults/defaults_Parameters.csv'

update_files = {}


### Required sector files (included in csv/model/)
sector_req = [
    'fuels',
    'transmission',
    'sector_coal mining',
    'sector_natural gas production',
    'sector_petroleum crude',
    'sector_petroleum refining',
    'sector_electricity',
    'sector_biodiesel',
    'sector_ethanol',
    'sector_hydrogen',
    'sector_mining',
    'sector_industrial minerals',
    'sector_iron and steel',
    'sector_metal smelting',
    'sector_chemical products',
    'sector_pulp and paper',
    'sector_light industrial',
    'sector_construction',
    'sector_residential',
    'sector_commercial',
    'sector_transportation personal',
    'sector_transportation freight',
    'sector_waste',
    'sector_agriculture',
    'sector_forestry',]
if sector_req:
    if model_path not in update_files:
        update_files[model_path] = []
    update_files[model_path].extend(sector_req)


### Required model files (included in csv/model/)
model_req = [
    'CIMS_DCC', # declining capital cost
    'CIMS_DIC', # declining intangible cost (neighbour effect)
    'CIMS_FIC', # fixed intangible cost; primarily used for calibration
    'CIMS_market share limits', # use of limits should be minimised
    ]
if model_req:
    if model_path not in update_files:
        update_files[model_path] = []
    update_files[model_path].extend(model_req)


### Optional model files (included in csv/model/)
model_optional = [
    'CIMS_exogenous prices',  # needed for correct calibration of historical years
    'CIMS_exogenous demand',  # use this file when calibrating endogenous supply sectors
    # 'CIMS_macro',
    ]
if model_optional:
    if model_path not in update_files:
        update_files[model_path] = []
    update_files[model_path].extend(model_optional)


### Optional scenario update files

ref_path = 'csv/policies/reference'
ref_policies = [
### Economy
    'Ref_carbon tax',
    'Ref_OBPS',
### Coal Mining
### Natural Gas Production
### Petroleum Crude
### Mining
### Electricity
    'Ref_coal phase out',
    'Ref_nuclear ban',
    'Ref_nuclear decommission',
    'Ref_clean electricity',
### Biodiesel
### Ethanol
### Hydrogen
### Petroleum Refining
### Industrial Minerals
### Iron and Steel
### Metal Smelting
### Chemical Products
### Pulp and Paper
### Light Industrial
### Residential
    'Ref_incandescent phase out',
### Commercial
### Transportation Personal
    'Ref_LDV ZEV',
    'Ref_renewable content transport fuels',
### Transportation Freight
### Waste
    'Ref_waste methane large sites',
### Agriculture
    ]
if ref_policies:
    if ref_path not in update_files:
        update_files[ref_path] = []
    update_files[ref_path].extend(ref_policies)


# These files turn off existing Reference policies starting in the first model year
# Off files must be updated annually to reflect the last common year between a set of scenarios
turn_off_path = 'csv/policies/turn off'                                           
turn_off_policies = [
### Economy
    # 'Off_OBPS',
### Coal Mining
### Natural Gas Production
### Petroleum Crude
### Mining
### Electricity
### Biodiesel
### Ethanol
### Hydrogen
### Petroleum Refining
### Industrial Minerals
### Iron and Steel
### Metal Smelting
### Chemical Products
### Pulp and Paper
### Light Industrial
### Residential
### Commercial
### Transportation Personal
### Transportation Freight
### Waste
### Agriculture
    ]
if turn_off_policies:
    if turn_off_path not in update_files:
        update_files[turn_off_path] = []
    update_files[turn_off_path].extend(turn_off_policies)


scenario_path = 'csv/policies/net zero'    # Change this path as necessary based on current scenario
scenario_policies = [
### Economy
    # 'NZ_carbon tax',
### Coal Mining
### Natural Gas Production
    # 'NZ_natural gas production',
### Petroleum Crude
    # 'NZ_petroleum crude',
### Mining
### Electricity
    # 'NZ_electricity generation',
### Biodiesel
### Ethanol
### Hydrogen
### Petroleum Refining
### Industrial Minerals
### Iron and Steel
### Metal Smelting
### Chemical Products
### Pulp and Paper
### Light Industrial
### Residential
### Commercial
### Transportation Personal
### Transportation Freight
### Waste
### Agriculture
    ]
if scenario_policies:
    if scenario_path not in update_files:
        update_files[scenario_path] = []
    update_files[scenario_path].extend(scenario_policies)



### Scenario Name
### This will be the save location for results (i.e., results_dir/scenario_name/results_general.csv)
scenario_name = 'Reference'  # Set this to current scenario (e.g., "Reference", "Net Zero")

# ---




# Leave the column list as-is to run the simulation
col_list1 = [
    'Branch',
    'Region',
    'Sector',
    'Technology',
    'Parameter',
    'Context',
    'Sub_Context',
    'Target',
    'Source',
    'Unit',
]
col_list2 = pd.Series(np.arange(2000,2051,5), dtype='string').tolist()
col_list = col_list1 + col_list2



# Base model and default values files below are required for the model to run
print(f'Loading base model: {base_model}')
load_paths = []
for reg in region_list:
    load_path = f'{model_path}/{base_model}/{base_model}_{reg}.csv'
    if path.exists(load_path):
        print(f'\t{reg} - loaded')
        load_paths.append(load_path)
    else:
        print(f'\t{reg} - file does not exist')

print(f'\nLoading model updates files:')
if bool(update_files):
    update_paths = []
    for dir, file in update_files.items():
        for file in update_files[dir]:
            print(f'\n{file}')
            for reg in region_list:
                update_path = f'{dir}/{file}/{file}_{reg}.csv'
                if path.exists(update_path):
                    print(f'\t{reg} - loaded')
                    update_paths.append(update_path)
                else:
                    print(f'\t{reg} - file does not exist')

else:
    print('None')



model = CIMS.Model(
        csv_init_file_paths = load_paths,
        csv_update_file_paths = update_paths,
        col_list = col_list,
        year_list = year_list,
        sector_list = sector_list,
        default_values_csv_path = default_path,
        )


# ::TODO:: Turn these back on!

#model.validate_files()
#model.build_graph()
#model.validate_graph()


model.run(equilibrium_threshold=0.05, max_iterations=10, show_warnings=False, print_eq=True)







## Below is just exporting and saving, same as before.







#################### Export results ###########################
results_path = f'results/{scenario_name}'
isExist = os.path.exists(results_path)
if not isExist:
   # Create a new directory because it does not exist
   os.makedirs(results_path)

CIMS.log_model(model=model,
                output_file = f'{results_path}/results_general.csv',
                path = 'results/results_general.txt')

print('\n')
print(f"Results exported to '{scenario_name}'")

### Export tech results ###
import os

results_path = f'results/{scenario_name}'
isExist = os.path.exists(results_path)
if not isExist:
   # Create a new directory because it does not exist
   os.makedirs(results_path)

CIMS.log_model(model=model,
                output_file = f'{results_path}/results_tech.csv',
                path = 'results/results_tech.txt')
print('\n')
print(f"Tech results exported to '{scenario_name}'")


