import os
import glob

# pretty much just removes everything from the folders, including ignoreMe.txt
def clear():
    path = os.getcwd()
    
    dts_input_path = path + "/DatavyuToSuperCoder/Input/"
    os.chdir(dts_input_path)
    dts_input_files = glob.glob(os.path.join(dts_input_path, "*"))
    for file in dts_input_files:
        os.remove(file)
        
    dts_output_path = path + "/DatavyuToSuperCoder/Output/"
    os.chdir(dts_output_path)
    dts_output_files = glob.glob(os.path.join(dts_output_path, "*"))
    for file in dts_output_files:
        os.remove(file)
        
    ft_path = path + "/combining/Input/"
    os.chdir(ft_path)
    ft_files = glob.glob(os.path.join(ft_path, "*"))
    for file in ft_files:
        os.remove(file)
        
    output_path = path + "/OUTPUT/"
    os.chdir(output_path)
    output_files = glob.glob(os.path.join(output_path, "*"))
    for file in output_files:
        os.remove(file)
        
    input_path = path + "/INPUT/"
    os.chdir(input_path)
    # only removes ignoreMe.txt from the INPUT folder (the .csv inputs stay)
    input_ignore = glob.glob(os.path.join(input_path, "*txt"))
    if input_ignore:
        os.remove(input_ignore[0])

clear()
