# MAIN IMPORT SECTION
import os

main_directory = input("Provide path where the files need to be saved: ")
variable = input("Input variable (precip/tmax/tmin): ")


# DEFINING FUNCTIONS
def working_directory(var):
    """determines the proper working directory for the variable in question"""
    import os
    parent_dir = main_directory
    var_dir = str(var)
    working_dir = os.path.join(parent_dir, var_dir)
    return working_dir


def clip_netcdf(var, year, shapefile_path, output_path, long_i=(-180), long_e=180, lat_i=(-50), lat_e=50):
    """clips the netCDF files to the desired shapefile"""
    # import section
    import os
    import xarray as xr
    import numpy as np
    import matplotlib.pyplot as plt
    import geopandas as gpd
    import regionmask
    # input section
    input_nc_location = os.path.join(working_directory(var), 'a_netCDFdata')
    nc_file_name = f"chirps-v2.0.{year}.days_p05.nc"
    input_nc = xr.open_dataset(os.path.join(input_nc_location, nc_file_name))
    # input shapefile
    ws_pol = gpd.read_file(r"{0}".format(shapefile_path))
    input_nc = input_nc.sel(longitude=slice(long_i, long_e), latitude=slice(lat_i, lat_e))
    # mask operations
    ws_pol_mask = regionmask.Regions(np.array(ws_pol.geometry))
    mask = ws_pol_mask.mask(input_nc.isel(time=0), lat_name='latitude', lon_name='longitude')
    mask_output = input_nc.where(mask == 0)
    # setting output path
    mask_output.to_netcdf(output_path)


# CLIPPING SCRIPT

# clipping algorithm
need_clipping = input("Do you need to clip some files? (Type 'Y' or 'N'): ")
while need_clipping != 'Y' and need_clipping != 'N':
    need_clipping = input("Please respond with 'Y' or 'N' ")
if need_clipping == 'Y':
    # setting up output location
    selected_working_directory = working_directory(variable)
    os.chdir(selected_working_directory)
    os.mkdir("ClippedNetCDF_files")
    clipped_files_location = os.path.join(selected_working_directory, "ClippedNetCDF_files")

    start_year = input('Input starting year from which clipping is needed: ')
    end_year = input('Input last year for which clipping is needed: ')
    print("Input the longitude and latitude to the nearest bounding integer")
    ln_i = int(input("Starting longitude(optional): "))
    ln_e = int(input("Ending longitude(optional): "))
    lt_i = int(input("Starting latitude(optional): "))
    lt_e = int(input("Ending latitude(optional): "))
    shapefile = input('Input complete path of mask shapefile: ')
    i = int(start_year)
    while i < int(end_year) + 1:
        output_clip_name = f'{i}clipped'
        output_clip = os.path.join(clipped_files_location, output_clip_name)
        clip_netcdf(variable, i, shapefile, output_clip, long_i=ln_i, long_e=ln_e, lat_i=lt_i, lat_e=lt_e)
        i += 1
