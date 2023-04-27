# import section
import os
import netCDF4
import numpy as np
from netCDF4 import Dataset
import openpyxl as op
from openpyxl import Workbook, load_workbook
import pandas as pd

working_directory = input("Full path of working directory: ")
variable = input('Input variable(precip/tmin/tmax): ')

no_of_grids_in_zone = [106, 83, 135, 110, 48]  # TO BE ADJUSTED

start_year = int(input("Initial year for which netCDF file needs to be converted: "))
end_year = int(input('Final year for which netCDF files need to be converted: '))

lat_long_file = input("Name of file with Lat-Long data with extension (must be present in the working directory): ")
lat_lon_xlsx = load_workbook(os.path.join(working_directory, lat_long_file)).active

ncdf_files_loc = input('Input directory path where NetCDF files are located: ')

current_year = start_year
while int(current_year) < int(end_year) + 1:
    data = Dataset(os.path.join(ncdf_files_loc, f"{current_year}clipped"))
    os.makedirs(os.path.join(working_directory, r'{0}\{1}zones'.format(variable, current_year)))
    zone_table_out_location = os.path.join(working_directory, r'{0}\{1}zones'.format(variable, current_year))
    # defining variables
    lat = data.variables['latitude'][:]  # TO BE UPDATED
    lon = data.variables['longitude'][:]  # TO BE UPDATED
    time = data.variables['time']  # TO BE UPDATED

    precip = data.variables['precip']  # TO BE UPDATED

    zone = 0
    while zone < len(no_of_grids_in_zone):
        # CREATION OF EMPTY PANDAS DATAFRAME
        # asserting the dates
        starting_date = f"{str(current_year)}-01-01"
        ending_date = f'{str(current_year)}-12-31'
        date_range = pd.date_range(start=starting_date, end=ending_date)

        # creation of empty dataframe
        dataframe = pd.DataFrame(f'{variable}', columns=['Grids'], index=date_range)
        year_days = np.arange(0, time.size)  # 'np.arange' creates arrays with increment of 1

        grid = 0
        while grid < no_of_grids_in_zone[zone]:

            # PREPARATION FOR ACCESSING PRECIP VALUES
            # defining latitudes and longitudes from Excel sheet
            lat_point = lat_lon_xlsx[f'{chr(65 + 2*zone)}{3 + grid}'].value
            lon_point = lat_lon_xlsx[f'{chr(66 + 2*zone)}{3 + grid}'].value

            # squared difference to find the nearest available grid point in netCDF
            lat_sq_dist = (lat_point - lat) ** 2
            lon_sq_dist = (lon_point - lon) ** 2

            # to find the index for the minimum values of the axis (squared difference)
            lat_min_index = lat_sq_dist.argmin()
            lon_min_index = lon_sq_dist.argmin()

            # PUTTING PRECIPITATION VALUES IN PANDAS DATAFRAME
            # access daily precipitation values of the grid in a list and making a new dataframe
            rainfall_values = list()
            for day in year_days:
                rain = precip[day, lat_min_index, lon_min_index]
                rainfall_values.append(rain)
            rainfall_dataframe = pd.DataFrame({f"rain_g{grid+1}": rainfall_values}, index=date_range)

            # put values in the pandas dataframe
            dataframe = pd.concat([dataframe, rainfall_dataframe], axis=1)

            print(f"Grid {grid+1} {variable} data of Zone_{zone+1} for {current_year} successfully generated")

            grid += 1

        dataframe.to_excel(os.path.join(zone_table_out_location, f'zone_{zone+1}.xlsx'))

        print(f"Excel file for Zone {zone+1} for the year {current_year} successfully generated")

        zone += 1

    print(f"Zonal {variable} data for {current_year} generated in {zone_table_out_location}")
    current_year += 1

print("All done")
