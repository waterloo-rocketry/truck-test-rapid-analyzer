# NOTE: This file is to be run within the directory above truck-test-rapid-analyzer!
import os

# Import hell
try:
    from ..truck_test_rapid_analyzer import execute_complete_analysis
except:
    # Hack-ey workaround that allows this to run using eclipse IDE
    import os
    import sys

    parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, parent)

    from truck_test_rapid_analyzer import execute_complete_analysis

if __name__ == "__main__":

    config = {}
    config['USE_HARD_CODED_PATH'] = False
    config['HARD_CODED_PATH'] = ""
    config['HARD_CODED_OUTPUT'] = ""
    config['FILE_INPUT_HITLIST'] = True
    config['HITLIST_PATH'] = "tests/example_hitlist.txt"
    config['SINGLE_OUTPUT_FILE'] = True
    config['SINGLE_OUTPUT_FILE_PATH'] = "Example Complete Analysis Direct.xlsx"
    config['FILE_SUBDIRECTORY'] = "tests/example_data_directory"
    config['USE_CUSTOM_SHEETNAMES'] = True
    config['SINGLE_OUTPUT_SHEETNAMES_PATH'] = r"tests/example_hitlist_sheetnames.txt"
    config['CONDENSED_EXPORT_VERSION'] = True
    config['SUPPRESS_ALL_PRINTS'] = False

    execute_complete_analysis(config)
