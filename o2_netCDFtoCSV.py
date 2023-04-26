# import section
import os
import netCDF4
import numpy as np
from netCDF4 import Dataset
import openpyxl as op
from openpyxl import Workbook, load_workbook
import pandas as pd

no_of_grids_in_zone = [105, 82, 134, 109, 47]
year = 2007
# dataset
while year < 2023:
    data = Dataset(r"C:\Users\Lenovo\Desktop\Main Project\Z1. Data\precip\ClippedNetCDF_files\{0}clipped".format(year))

    # defining variables
    lat = data.variables['latitude'][:]
    lon = data.variables['longitude'][:]
    time = data.variables['time'][:]

    precip = data.variables['precip']

    lat_lon_xlsx = load_workbook(r"C:\Users\Lenovo\Desktop\Main Project\Z1. Data\Zone_latLon.xlsx").active

    zone = 0
    while zone < 5:
        grid = 0
        while grid < no_of_grids_in_zone[zone]:

            # PREPARATION FOR ACCESSING PRECIP VALUES
            # defining latitudes and longitudes from Excel sheet
            lat_point = lat_lon_xlsx[f'{chr(65 + zone)}{3 + grid}'].value
            lon_point = lat_lon_xlsx[f'{chr(66 + zone)}{3 + grid}'].value

            # squared difference to find the nearest available grid point in netCDF
            lat_sq_dist = (lat_point - lat)**2
            lon_sq_dist = (lon_point - lon)**2

            # to find the index for the minimum values of the axis (squared difference)
            lat_min_index = lat_sq_dist.argmin()
            lon_min_index = lon_sq_dist.argmin()

            # CREATION OF PANDAS DATAFRAME
            # asserting the dates
            starting_date = f"{str(year)}{data.variables['time'].units[15:27]}"
            ending_date = f'{str(year)}-12-31'
            date_range = pd.date_range(start=starting_date, end=ending_date)

            # creation of empty dataframe
            dataframe = pd.DataFrame(0, columns=['Rainfall'], index=date_range)
            year_days = np.arange(0, data.variables['time'].size)
            for day in year_days:
                dataframe.iloc[day] = precip[day, lat_min_index, lon_min_index]

            grid += 1
        zone += 1
    year += 1
