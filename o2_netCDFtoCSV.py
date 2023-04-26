# import section
import os
import netCDF4
import numpy as np
from netCDF4 import Dataset
import openpyxl as op
from openpyxl import Workbook, load_workbook
import pandas as pd

variable = input('Input variable(precip/tmin/tmax): ')
no_of_grids_in_zone = [105, 82, 134, 109, 47]
year = 2007
lat_lon_xlsx = load_workbook(r"C:\Users\Lenovo\Desktop\Main Project\Z1. Data\Zone_latLon.xlsx").active
# dataset
while year < 2023:
    data = Dataset(r"C:\Users\Lenovo\Desktop\Main Project\Z1. Data\{0}\ClippedNetCDF_files\{1}clipped".format(variable, year))
    os.makedirs(r'C:\Users\Lenovo\Desktop\Main Project\Z1. Data\{0}\{1}zones'.format(variable, year))
    zone_table_out_location = r'C:\Users\Lenovo\Desktop\Main Project\Z1. Data\{0}\{1}zones'.format(variable, year)
    # defining variables
    lat = data.variables['latitude'][:]
    lon = data.variables['longitude'][:]
    time = data.variables['time'][:]

    precip = data.variables['precip']

    zone = 0
    while zone < 5:
        # CREATION OF EMPTY PANDAS DATAFRAME
        # asserting the dates
        starting_date = f"{str(year)}-01-01"
        ending_date = f'{str(year)}-12-31'
        date_range = pd.date_range(start=starting_date, end=ending_date)

        # creation of empty dataframe
        dataframe = pd.DataFrame('Rainfall', columns=['Grids'], index=date_range)
        year_days = np.arange(0, data.variables['time'].size)  # 'np.arange' creates arrays with increment of 1

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

            print(f"Grid{grid+1}{variable} for {year} data successfully generated")

            grid += 1

        dataframe.to_excel(os.path.join(zone_table_out_location, f'zone_{zone+1}.xlsx'))

        print(f"Excel file for Zone{zone+1} successfully generated")

        zone += 1

    print(f"Zonal {variable} data for {year} generated in {zone_table_out_location}")
    year += 1
