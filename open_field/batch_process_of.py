

import os
import sys
import numpy as np
import pandas as pd

input_text = input('Provide the filepath of the files: ')

direc = input_text

while True:
    try:
        files = os.listdir(direc)
        excel_files = []

        for file in files:
            if '.xlsx' in file:
                excel_files = np.append(file, excel_files)

        while True:
            if len(excel_files) > 1:
                print('Working...')

                final_df = pd.DataFrame(columns = ['Avg Distance Moved (cm)', 'Avg Velocity (cm/s)', 'Percent in Zone', 'ID', 'Sex', 'Tx'])

                for file in excel_files:

                    print('Processing ', file)

                    ## use the special excel pandas function to read in the .xlsx file (this allows us to access the sheets individually)
                    xls = pd.ExcelFile(direc + '\\' + file)
                    
                    ## i can use this method to retrieve the list of sheet names! this will be useful for iterating through
                    sheet_names = xls.sheet_names
                    
                    ## create the empty output dataframe that will soon contain all our desired information!
                    output_df = pd.DataFrame(columns = ['Avg Distance Moved (cm)', 'Avg Velocity (cm/s)', 'Percent in Zone', 'ID', 'Sex', 'Tx'])

                    for sheet in sheet_names:

                        df = pd.read_excel(xls, sheet, header = 38)
                        df.drop(0, axis = 0, inplace = True)
                        df.replace('-', 0, inplace = True)

                        ## collecting the metadata output from the Ethovision data table

                        df_meta = pd.read_excel(xls, sheet, index_col = 0)
                        metadata = df_meta.loc[['Animal ID', 'Sex', 'Condition']].dropna(axis = 1).T.values

                        print('Sheet metadata: ', metadata)
                        if len(metadata)>0:
                            ## my goal is to take the raw data from the open field tests and analzye each one by first one minute, 
                            ## first 2 minutes, and first 5 minutes (ethovision spits out just averages for the whole 10 min

                            ## to keep things clean, I am going to start off by grabbing the index locations for the three time points
                            one_min = df[df['Recording time'] == 60].index[0]
                            two_min = df[df['Recording time'] == 120].index[0]
                            five_min = df[df['Recording time'] == 300].index[0]

                            ## i just want total distance moved, 

                            distance_moved = []

                            distance_moved.append(df['Distance moved'].values[0:one_min].mean())
                            distance_moved.append(df['Distance moved'].values[0:two_min].mean())
                            distance_moved.append(df['Distance moved'].values[0:five_min].mean())

                            ## average velocity,

                            velocity = []

                            velocity.append(df['Velocity'].values[0:one_min].mean())
                            velocity.append(df['Velocity'].values[0:two_min].mean())
                            velocity.append(df['Velocity'].values[0:five_min].mean())

                            ## and maybe percent of time in zone 1 (which is the center)

                            zone_1 = []

                            zone_1.append(df['In zone'].values[0:one_min].sum()/one_min)
                            zone_1.append(df['In zone'].values[0:two_min].sum()/two_min)
                            zone_1.append(df['In zone'].values[0:five_min].sum()/five_min)

                            mini_df = pd.DataFrame([distance_moved, velocity, zone_1], index = ['Avg Distance Moved (cm)', 'Avg Velocity (cm/s)', 'Percent in Zone'], 
                                                columns = ['1 Min', '2 Min', '5 Min']).T

                            mini_df['ID'] = metadata[0][0]
                            mini_df['Sex'] = metadata[0][1]
                            mini_df['Tx'] = metadata[0][2]

                            output_df = output_df.append(mini_df)
                        
                    final_df = final_df.append(output_df)


                final_df.to_csv(direc + '\\' + 'open_field_output.csv')

                print('Output file is located here: ', direc + '\\' + 'open_field_output.csv')

            if len(excel_files) < 1:
                print('Maybe try a different folder? The one provided did not have any output sheets')
                exit()
            if direc + '\\' + 'open_field_output.csv':
                quit()


    except:
        if direc + '\\' + 'open_field_output.csv':
            quit()

