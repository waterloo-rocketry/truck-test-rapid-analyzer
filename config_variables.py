# Configuration Variables:

# Configures the script to analyze one specific data file.
# This option is overridden if a file input hitlist is used (see further options below)
USE_HARD_CODED_PATH = False

# Input path for hard-coded option
HARD_CODED_PATH = r"C:\example\given\path\a_file.csv"
HARD_CODED_OUTPUT = "example_output.xlsx"  # Output path for hard-coded option


FILE_INPUT_HITLIST = True  # If the script should input a list of file paths as input
HITLIST_PATH = r"default_target_list.txt"  # Location of list

# If the analysis should be output as one file (recommended)
SINGLE_OUTPUT_FILE = True
SINGLE_OUTPUT_FILE_PATH = "Complete Analysis.xlsx"  # Name of single file

# If there is a seperate folder containing all inputs and outputs, it is specified here
FILE_SUBDIRECTORY = r"default_data_folder"


# Whether or not a custom name should be used for the sheetname
USE_CUSTOM_SHEETNAMES = True

# The location of the file where an ordered list of a custom name for every file
# on the input hitlist is provided. This is ignored if custom sheenames are not in use
SINGLE_OUTPUT_SHEETNAMES_PATH = r"default_target_list_sheetnames.txt"


# When the analysis is output as a single file, this variable controls whether
# the Data and Graphs for each csv data file will be in two seperate labelled sheets (if False),
# or together in one sheet with both the data and grpahs of that data file (if True)
CONDENSED_EXPORT_VERSION = True

# For controlling printouts to console
DEBUG_MODE = True
DEBUG_MODE_VERBOSE = False

# End of configuration variables
