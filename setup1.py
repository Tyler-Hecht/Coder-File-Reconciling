import os
import glob
import shutil
import subprocess
from openpyxl import *
del open
from pandas import *

path = os.getcwd()
if os.name == "nt":
    sep = "\\"
else:
    sep = "/"

input_path = path + "/INPUT/"
output_path = path + "/OUTPUT/"
combined_file = glob.glob(os.path.join(input_path, "*.xlsx"))
third_code = glob.glob(os.path.join(input_path, "*.csv"))

# get input files
if len(combined_file) == 0:
    print("ERROR: no combined file found")
    exit(1)
if len(third_code) == 0:
    print("ERROR: no third coder file found")
    exit(1)
if len(combined_file) > 2:
    print("ERROR: multiple combined files found")
    exit(1)
if len(third_code) > 2:
    print("ERROR: multiple third coder files found")
    exit(1)
combined_file = combined_file[0]
combined_file_name = combined_file.split(sep)[-1]
third_code = third_code[0]
third_code_name = third_code.split(sep)[-1]

os.chdir(input_path)
dts_input_path = path + "/DatavyuToSupercoder/Input/"
dts_output_path = path + "/DatavyuToSupercoder/Output/"

# move file to DTS input
shutil.copyfile(third_code_name, dts_input_path + third_code_name)

