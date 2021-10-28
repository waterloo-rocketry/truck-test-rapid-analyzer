from openpyxl import Workbook as WB
from openpyxl import load_workbook as load_wb

from openpyxl.utils import get_column_letter
from openpyxl.utils import absolute_coordinate as abs_coords

from openpyxl.workbook.defined_name import DefinedName as DName

from openpyxl.chart import LineChart, ScatterChart, Reference, Series

import csv
from openpyxl.styles.alignment import Alignment

# Configuration Variables:

USE_HARD_CODED_PATH = True # Configures the script to analyze one specific data file
HARD_CODED_PATH = r"C:\example\given\path\a_file.csv" # Input path for this option
HARD_CODED_OUTPUT = "example_output.xlsx" # Output path for this option

FILE_INPUT_HITLIST = True # If the script should input a list of file paths as input
HITLIST_PATH = r"target_list.txt" # Location of list

SINGLE_OUTPUT_FILE = True # If the analysis should be input as one file
SINGLE_OUTPUT_FILE_PATH = "Complete Analysis.xlsx"
SINGLE_OUTPUT_SHEETNAMES_PATH = r"target_list_sheetnames.txt"

CONDENSED_EXPORT_VERSION = True

DEBUG_MODE = True
DEBUG_MODE_VERBOSE = False


def space_columns(target_sheet, depth=1, cols=0):
    '''
    Space the columns to make sure all text in headings is visible.
    
    Parameters
    ----------
    target_sheet: openpyxl.Worksheet
        The sheet where the columns must be spaced.
    depth: int=1
        How deep the length check should be done, for vertical headers. Default value is 1
    cols: int=0
        How many columns should be checked. Default values is 0, which specifies a scan of all
        columns.
    '''
    
    if cols:
        for col_count in range(1, cols + 1):
            max_width = 0
            for didx in range(1, depth + 1):
                if isinstance(target_sheet.cell(didx, col_count).value, str):
                    curr_width = len(target_sheet.cell(didx, col_count).value)
                    
                    target_sheet.cell(didx, col_count).alignment = Alignment(wrap_text=True)
                    
                    if curr_width > max_width:
                        max_width = curr_width

            target_sheet.column_dimensions[get_column_letter(
                col_count)].width = int(0.75 * max_width)
            target_sheet.cell(1, col_count).alignment = Alignment(wrap_text=True)

    else:  # Scan columns until they are empty

        col_count = 1
        while target_sheet.cell(1, col_count).value:
            max_width = 0
            for didx in range(1, depth + 1):
                if isinstance(target_sheet.cell(didx, col_count).value, str):
                    curr_width = len(target_sheet.cell(didx, col_count).value)
                    
                    target_sheet.cell(didx, col_count).alignment = Alignment(wrap_text=True)
                    
                    if curr_width > max_width:
                        max_width = curr_width

            target_sheet.column_dimensions[get_column_letter(
                col_count)].width = int(0.75 * max_width)
            target_sheet.cell(1, col_count).alignment = Alignment(wrap_text=True)
            
            col_count += 1


class Consts():
    """
    Data and utility class to store constants data.
    """

    L_NOM_DIAM_TITLE = "A1"

    L_NOM_DIAM_M = "B1"
    NOM_DIAM_M = "Constants!" + abs_coords(L_NOM_DIAM_M)

    L_NOM_DIAM_FT = "C1"
    NOM_DIAM_FT = "Constants!" + abs_coords(L_NOM_DIAM_FT)


    L_AIR_DENSITY_TITLE = "A2"

    L_AIR_DENSITY_KG_M3 = 'B2'
    AIR_DENSITY_KG_M3 = "Constants!" + abs_coords(L_AIR_DENSITY_KG_M3)

    L_AIR_DENSITY_SLG_FT3 = 'C2'
    AIR_DENSITY_SLG_FT3 = "Constants!" + abs_coords(L_AIR_DENSITY_SLG_FT3)


    L_DESCENT_RATE_TITLE = "A3"

    L_DESCENT_RATE_FT_S = "B3"
    DESCENT_RATE_FT_S = "Constants!" + abs_coords(L_DESCENT_RATE_FT_S)

    L_DESCENT_RATE_M_S = "C3"
    DESCENT_RATE_M_S = "Constants!" + abs_coords(L_DESCENT_RATE_M_S)


    L_ROCKET_MASS_TITLE = 'A4'

    L_ROCKET_MASS_LB = "B4"
    ROCKET_MASS_LB = "Constants!" + abs_coords(L_ROCKET_MASS_LB)

    L_ROCKET_MASS_KG = "C4"
    ROCKET_MASS_KG = "Constants!" + abs_coords(L_ROCKET_MASS_KG)


    L_NOM_SA_TITLE = 'A5'

    L_NOM_SA_M2 = "B5"
    NOM_SA_M2 = "Constants!" + abs_coords(L_NOM_SA_M2)

    L_NOM_SA_FT2 = "C5"
    NOM_SA_FT2 = "Constants!" + abs_coords(L_NOM_SA_FT2)


    L_TARGET_DRAG_AREA_TITLE = "A6"

    L_TARGET_DRAG_AREA_M2 = "B6"
    TARGET_DRAG_AREA_M2 = "Constants!" + abs_coords(L_TARGET_DRAG_AREA_M2)

    L_TARGET_DRAG_AREA_FT2 = "C6"
    TARGET_DRAG_AREA_FT2 = "Constants!" + abs_coords(L_TARGET_DRAG_AREA_FT2)


    L_ANEMOMETER_FACTOR_TITLE = "A7"

    L_ANEMOMETER_FACTOR = "B7"
    ANEMOMETER_FACTOR = "Constants!" + abs_coords(L_ANEMOMETER_FACTOR)


    L_LOAD_CELL_FACTOR_TITLE = 'A8'

    L_LOAD_CELL_FACTOR = 'B8'
    LOAD_CELL_FACTOR = "Constants!" + abs_coords(L_LOAD_CELL_FACTOR)


    L_TARGET_COEFF_DRAG_TITLE = 'A9'

    L_TARGET_COEFF_DRAG = 'B9'
    TARGET_COEFF_DRAG = "Constants!" + abs_coords(L_TARGET_COEFF_DRAG)

    def create_defined_names(self, target_workbook):
        """
        Turn the constants into excel defined names.
        
        Parameters
        ----------
            target_workbook: openpyxl.Workbook
                The workbook where the defined names are the be created.
        """
        names = target_workbook.defined_names # Used to shorten lines 
        
        names.append(DName('NOM_DIAM_M', attr_text=self.NOM_DIAM_M))
        names.append(DName('NOM_DIAM_FT', attr_text=self.NOM_DIAM_FT))

        names.append(DName('AIR_DENSITY_KG_M3', attr_text=self.AIR_DENSITY_KG_M3))
        names.append(DName('AIR_DENSITY_SLG_FT3', attr_text=self.AIR_DENSITY_SLG_FT3))

        names.append(DName('DESCENT_RATE_M_S', attr_text=self.DESCENT_RATE_M_S))
        names.append(DName('DESCENT_RATE_FT_S', attr_text=self.DESCENT_RATE_FT_S))

        names.append(DName('ROCKET_MASS_LB', attr_text=self.ROCKET_MASS_LB))
        names.append(DName('ROCKET_MASS_KG', attr_text=self.ROCKET_MASS_KG))

        names.append(DName('TARGET_DRAG_AREA_M2', attr_text=self.TARGET_DRAG_AREA_M2))
        names.append(DName('TARGET_DRAG_AREA_FT2', attr_text=self.TARGET_DRAG_AREA_FT2))

        names.append(DName('NOM_SA_M2', attr_text=self.NOM_SA_M2))
        names.append(DName('NOM_SA_FT2', attr_text=self.NOM_SA_FT2))

        names.append(DName('ANEMOMETER_FACTOR', attr_text=self.ANEMOMETER_FACTOR))

        names.append(DName('LOAD_CELL_FACTOR', attr_text=self.LOAD_CELL_FACTOR))

        names.append(DName('TARGET_COEFF_DRAG', attr_text=self.TARGET_COEFF_DRAG))

def f_create_constants_table(target_sheet, loc=None):
    """
    Create the constants table for aerodynamic analysis.
    
    Parameters
    ----------
    target_sheet: openpyxl.Worksheet
        The sheet where the constants table should be created.
    loc:
        Locations of the constants, if specifying it is necessary. Currently implemnted
        as a global class, but in the future it may be desired to implement it as an
        argument.
    """

    target_sheet[Consts.L_NOM_DIAM_TITLE] = "Nominal diamater of parachute (m | ft)"
    target_sheet[Consts.L_NOM_DIAM_M] = 4.965
    target_sheet[Consts.L_NOM_DIAM_FT] = "=B1/0.3048"

    target_sheet[Consts.L_AIR_DENSITY_TITLE] = "Air density (kg/m^3 | slug/ft^3) "
    target_sheet[Consts.L_AIR_DENSITY_KG_M3] = 1.225
    target_sheet[Consts.L_AIR_DENSITY_SLG_FT3] = "=B2*0.00194032"

    target_sheet[Consts.L_DESCENT_RATE_TITLE] = "target descent rate (ft | m)"
    target_sheet[Consts.L_DESCENT_RATE_FT_S] = 112
    target_sheet[Consts.L_DESCENT_RATE_M_S] = "=B3*0.3048"

    target_sheet[Consts.L_ROCKET_MASS_TITLE] = "Rocket Mass (lb | kg)"
    target_sheet[Consts.L_ROCKET_MASS_LB] = 100
    target_sheet[Consts.L_ROCKET_MASS_KG] = "=B4*0.453"

    target_sheet[Consts.L_NOM_SA_TITLE] = "Nominal surface area of parachute (m^2 | ft^2)"
    target_sheet[Consts.L_NOM_SA_M2] = "=B1^2 * PI() * 0.25"
    target_sheet[Consts.L_NOM_SA_FT2] = "=C1^2 * PI() * 0.25"

    target_sheet[Consts.L_TARGET_DRAG_AREA_TITLE] = "Target Drag Area (m^2 | ft^2)"
    target_sheet[Consts.L_TARGET_DRAG_AREA_M2] = "=(2*C4)/(C3^2 * B2)"
    target_sheet[Consts.L_TARGET_DRAG_AREA_FT2] = "=(2*B4)/(B3^2 * C2)"

    target_sheet[Consts.L_ANEMOMETER_FACTOR_TITLE] = \
        "    Anemometer Adjustment (Calibration) Factor (unitless)"
    target_sheet[Consts.L_ANEMOMETER_FACTOR] = 0.725

    target_sheet[Consts.L_LOAD_CELL_FACTOR_TITLE] = \
            "Load Cell Adjustment (Calibration) Factor (unitless)"
    target_sheet[Consts.L_LOAD_CELL_FACTOR] = 1.0

    target_sheet[Consts.L_TARGET_COEFF_DRAG_TITLE] = \
            "Target Coeffcient of Drag Relative to Nominal SA (unitless)"
    target_sheet[Consts.L_TARGET_COEFF_DRAG] = "=B6/B5"
    
    


def create_headers(sheet, target_row=1):
    """
    Create headers in specified sheet.
    
    Parameters
    ----------
    sheet: openpyxl.Worksheet
        The sheet where the headers are to be created

    target_row: The row where the headers should be inserted. Default is 1, which is the first row.
                (openpyxl starts counting at 1 so this is the true first row in the spreadsheet)
    
    """
    sheet.cell(target_row, 1, "Time (ms)")
    sheet.cell(target_row, 2, "Anemometer Raw (m/s)")
    sheet.cell(target_row, 3, "Load Cell Raw (lbf)")
    sheet.cell(target_row, 4, "Anemometer Calibrated (m/s)")
    sheet.cell(target_row, 5, "Load Cell Calibrated (lbf)")

    sheet.cell(target_row, 6, "Load Cell Averaged (lbf)")
    sheet.cell(target_row, 7, "Anemometer Averaged (m/s)")
    sheet.cell(target_row, 8, "Anemometer (ft/s)")

    sheet.cell(target_row, 9, "Target Force (lbf)")
    sheet.cell(target_row, 10, "Drag Area [CdSo] (ft^2)")
    sheet.cell(target_row, 11, "Drag Coefficient (unitless)")


def process_sheet_row(raw_data, row_idx, sheet):
    """
    Create all formulas and input sheet data for a single data timestamp (row)
    
    Parameters
    ----------
    raw_data: array of float-like str
        The raw data from the DAQ csv
    row_idx: int
        The index of the row where it is the to be inserted. That is, the target excel sheet row
    sheet: openpyxl.Worksheet
        The worksheet where the data will be inputted.
    """
    sheet.cell(row_idx, 1, int(float(raw_data[0])*1000))

    sheet.cell(row_idx, 2, float(raw_data[1]))
    sheet.cell(row_idx, 2).number_format = "0.000"

    sheet.cell(row_idx, 3, float(raw_data[2]))
    sheet.cell(row_idx, 3).number_format = "0.000"

    r_force = "C" + str(row_idx)
    r_wspeed = "B" + str(row_idx)

    calibrated_windspeed_formula = "=" +\
        r_wspeed + "/" + 'ANEMOMETER_FACTOR'
    sheet.cell(row_idx, 4, calibrated_windspeed_formula)
    sheet.cell(row_idx, 4).number_format = "0.000"
    cal_ws = "D" + str(row_idx)

    calibrated_force_formula = "=" + r_force + \
        "/" + 'LOAD_CELL_FACTOR'
    sheet.cell(row_idx, 5, calibrated_force_formula)
    sheet.cell(row_idx, 5).number_format = "0.000"
    cal_force = "E" + str(row_idx)

    force_average_window = 3
    if row_idx - force_average_window >= 1:
        averaged_force_lower_bound = row_idx - force_average_window
    else:
        averaged_force_lower_bound = 1

    averaged_force_upper_bound = row_idx + force_average_window

    f_bounds = cal_force[0] + str(averaged_force_lower_bound) + ":" +\
        cal_force[0] + str(averaged_force_upper_bound)

    averaged_force_formula = "=AVERAGE(" + f_bounds + ")"
    sheet.cell(row_idx, 6, averaged_force_formula)
    sheet.cell(row_idx, 6).number_format = "0.000"
    av_force = "F" + str(row_idx)

    #

    ws_average_window = 0
    if row_idx - ws_average_window >= 1:
        averaged_ws_lower_bound = row_idx - ws_average_window
    else:
        averaged_ws_lower_bound = 1

    averaged_ws_upper_bound = row_idx + ws_average_window

    ws_bounds = cal_ws[0] + str(averaged_ws_lower_bound) + ":" +\
        cal_ws[0] + str(averaged_ws_upper_bound)

    averaged_windspeed_formula = "=AVERAGE(" + ws_bounds + ")"
    sheet.cell(row_idx, 7, averaged_windspeed_formula)
    sheet.cell(row_idx, 7).number_format = "0.000"
    av_windspeed = "G" + str(row_idx)

    ft_s_windspeed_formula = "=" + av_windspeed + "/0.3048"
    sheet.cell(row_idx, 8, ft_s_windspeed_formula)
    sheet.cell(row_idx, 8).number_format = "0.000"
    av_windspeed_ft_s = "H" + str(row_idx)

    target_force_formula = "=(" + av_windspeed_ft_s +\
        "^2)*" + 'AIR_DENSITY_SLG_FT3' +\
        "*" + 'TARGET_DRAG_AREA_FT2' +\
        "*0.5"

    sheet.cell(row_idx, 9, target_force_formula)
    sheet.cell(row_idx, 9).number_format = "0.000"
    target_force_lbf = "I" + str(row_idx)

    drag_area_formula = "=if(" + av_windspeed_ft_s + "=0, ," +\
        "(2*" + av_force + ")/(" + 'AIR_DENSITY_SLG_FT3' +\
        "*(" + av_windspeed_ft_s + ")^2))"
    sheet.cell(row_idx, 10, drag_area_formula)
    sheet.cell(row_idx, 10).number_format = "0.000"
    drag_area_ft2 = "J" + str(row_idx)

    coeff_of_drag_formula = "=" + drag_area_ft2 +\
        "/" + 'NOM_SA_FT2'
    sheet.cell(row_idx, 11, coeff_of_drag_formula)
    sheet.cell(row_idx, 11).number_format = "0.000"
    coeff_of_drag = "K" + str(row_idx)


def create_graphs(data_sheet, output_sheet, max_idx):
    """
    Create graphs from the data.
    
    The graphs unfortunately do not transfer to google sheets formats, so they will need to be
    re-created once in google sheets.
    
    Parameters
    ----------
    data_sheet: openpyxl.Worksheet
        The sheet where the data is located
    output_sheet: openpyxl.Worksheet
        The sheet where the graph should be located
    max_idx: int
        The maximum shreadsheet row where the data is
    """
    
    chart = ScatterChart()
    drag_area_chart = ScatterChart()
    #chart.style = 13
    time_data = Reference(data_sheet, min_col=1, min_row=2, max_row=max_idx)
    wind_data = Reference(data_sheet, min_col=2, min_row=2, max_row=max_idx)
    force_data = Reference(data_sheet, min_col=3, min_row=2, max_row=max_idx)
    cal_wind_data = Reference(data_sheet, min_col=4,
                              min_row=2, max_row=max_idx)
    av_force_data = Reference(data_sheet, min_col=6,
                              min_row=2, max_row=max_idx)
    target_force_data = Reference(
        data_sheet, min_col=9, min_row=2, max_row=max_idx)
    drag_area_data = Reference(
        data_sheet, min_col=10, min_row=2, max_row=max_idx)

    wind_series = Series(wind_data, xvalues=time_data, title="Raw Windspeed")
    force_series = Series(force_data, xvalues=time_data, title="Raw Force")
    cal_wind_series = Series(cal_wind_data, xvalues=time_data,
                             title="Calibrated Windspeed")
    av_force_series = Series(av_force_data, xvalues=time_data,
                             title="Averaged Force")
    target_force_series = Series(target_force_data, xvalues=time_data,
                                 title="Target Force")
    drag_area_series = Series(drag_area_data, xvalues=time_data,
                              title="Drag Area")

    chart.append(wind_series)
    chart.append(force_series)
    chart.append(cal_wind_series)
    chart.append(av_force_series)
    chart.append(target_force_series)

    drag_area_chart.append(drag_area_series)

    if DEBUG_MODE_VERBOSE:
        print(data_sheet.cell(max_idx - 1, 1).value)

    chart.x_axis.scaling.max = data_sheet.cell(max_idx - 1, 1).value
    chart.x_axis.scaling.min = 0
    chart.height = 10
    chart.width = 30
    chart.x_axis.title = "Time"

    chart.y_axis.axId = 200

    drag_area_chart.x_axis.scaling.max = data_sheet.cell(max_idx - 1, 1).value
    drag_area_chart.x_axis.scaling.min = 0
    drag_area_chart.height = 10
    drag_area_chart.width = 40
    drag_area_chart.y_axis.title = "Drag Area Axis"
    drag_area_chart.y_axis.majorGridlines = None

    drag_area_chart.y_axis.crosses = "max"

    drag_area_chart += chart
    if output_sheet == None:
        output_sheet = data_sheet
        output_sheet.add_chart(drag_area_chart, 'M2')
    else:
        output_sheet.add_chart(drag_area_chart, 'B2')

def simple_filter(raw_data): # to be expanded later
    """
    Dummy filter function for sifting through data, currently just returns True.
    
    Parameters
    ----------
    raw_data: array if float-like str
        The data based on which the filtering will occur.
    """
    return True

def execute_analysis(input_path, output_path, 
                     single_file=False, 
                     single_workbook=None, 
                     sheetname_prefix=None, 
                     condensed_version=False):
    """
    Execute a full analysis of a single data set.
    
    input_path: path-format str
        The path where the csv data is located
    output_path: path-format str
        The path where the excel file should be saved
    single_file: bool
        Whether the analysis should be performed as part of a single output sheet.
        Defaults to False.
    sheetname_prefix: str  
        In the event of a single file output, what the prefix of the sheet would be.
    condensed_version: bool
        In the event of a single file output, whether it should be done in a condensed
        output version. This is mostly used for the google drive output.
    """

    if single_workbook:
        
        
        if not condensed_version:
            ws_data = single_workbook.create_sheet(title=("Data " + sheetname_prefix))
            ws_graphs = single_workbook.create_sheet(title=("Graphs " + sheetname_prefix))
            ws_data.sheet_view.zoomScale = 70
        else:
            ws_data = single_workbook.create_sheet(title=(sheetname_prefix))
            ws_data.sheet_view.zoomScale = 55
    else:
        active_wb = WB()
        ws_data = active_wb.create_sheet(title="Data")
        ws_consts = active_wb.create_sheet(title="Constants")
        ws_graphs = active_wb.create_sheet(title="Graphs")

        f_create_constants_table(ws_consts)
        Consts.create_defined_names(active_wb)

    create_headers(ws_data)

    csvfile = open(input_path + '.csv', newline='')
    csvreader = csv.reader(csvfile, delimiter=',')

    row_idx = 1
    for row in list(csvreader):
        if row_idx and row_idx != 1:
            if simple_filter(row):  # Data Filtering
                process_sheet_row(row, row_idx, ws_data)
            else:
                row_idx -= 1

        row_idx += 1
    
    if condensed_version:
        create_graphs(ws_data, None, row_idx)
    else:
        create_graphs(ws_data, ws_graphs, row_idx)

    space_columns(ws_data, 1)

    if not single_workbook:
        space_columns(ws_consts, 10, 1)
        active_wb.save(output_path)



if __name__ == "__main__":
    print('Analyzer active')
    
    if FILE_INPUT_HITLIST:
        hitlist_path = HITLIST_PATH
        hitlist = []
        with open(hitlist_path, 'r') as hitlist_file:
            hitlist = list(hitlist_file)

        prefixes_path = SINGLE_OUTPUT_SHEETNAMES_PATH
        prefixes = []
        with open(prefixes_path, 'r') as prefixes_file:
            prefixes = list(prefixes_file)


        if SINGLE_OUTPUT_FILE:
            print('Exporting analysis to single file: '
                  + SINGLE_OUTPUT_FILE_PATH)

            global_workbook = WB()
            ws_consts = global_workbook.create_sheet(title="Constants")
            f_create_constants_table(ws_consts)
            Consts().create_defined_names(global_workbook)
            space_columns(ws_consts, 10, 1)


            for target, prefix in zip(hitlist, prefixes):
                execute_analysis(target[:-1],
                                 output_path=target[:-1] + '___analyzed.xlsx',
                                 single_file=True,
                                 single_workbook=global_workbook,
                                 sheetname_prefix=prefix,
                                 condensed_version=CONDENSED_EXPORT_VERSION)

                print(target[:-1] + " analyzed")


            global_workbook.save(SINGLE_OUTPUT_FILE_PATH)
            print('\n Analysis complete')
        else:
            for target in hitlist:
                execute_analysis(target[:-1],
                                 output_path=target[:-1] + '___analyzed.xlsx')

                print(target[:-1] + " analyzed")

            print('\n Analysis complete')
    else:
        if USE_HARD_CODED_PATH:
            itarget = HARD_CODED_PATH
            otarget = HARD_CODED_OUTPUT
        else:
            itarget = input("Enter the input target path")
            otarget = input("Enter the output target path")

        active_wb = WB()
        ws_data = active_wb.create_sheet(title="Data")
        ws_consts = active_wb.create_sheet(title="Constants")
        ws_graphs = active_wb.create_sheet(title="Graphs")

        f_create_constants_table(ws_consts)
        create_headers(ws_data)

        csvfile = open(itarget, newline='')
        csvreader = csv.reader(csvfile, delimiter=',')

        row_idx = 1
        for row in list(csvreader):
            if row_idx and row_idx != 1:
                if simple_filter(row):  # Data Filtering
                    process_sheet_row(row, row_idx, ws_data)
                else:
                    row_idx -= 1

            row_idx += 1

        create_graphs(ws_data, ws_graphs, row_idx)

        space_columns(ws_data, 1)
        space_columns(ws_consts, 10, 1)

        active_wb.save(otarget)
