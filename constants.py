from openpyxl import Workbook as WB
from openpyxl import load_workbook as load_wb

from openpyxl.utils import get_column_letter
from openpyxl.utils import quote_sheetname as quote
from openpyxl.utils import absolute_coordinate as abs_coords

from openpyxl.workbook.defined_name import DefinedName as DName

from openpyxl.chart import LineChart, ScatterChart, Reference, Series


class SimplifiedConsts():
    """
    Data and utility class to store constants data. 

    All of the constants and the class are treated in a more object-oriented fashion.
    An advantage of doing this is that all locations will be determined dymanically based on
    what order thay are added in, which means that adding and removing variables will be much
    easier. 

    """

    def __init__(self, target_worksheet=None, target_workbook=None):
        self.target_workbook = target_workbook
        self.target_worksheet = target_worksheet
        self._next_row_idx = 1

        self.variables_raw = []
        self.defined_name_dict = {}

        self.add_all_fields()

    def add_all_fields(self):
        """
        Adds all the variable fields required in the calculator
        """
        self.add_variable_field("Nominal diamater of parachute (m | ft)",
                                "NOM_DIAM_M", 4.965, "NOM_DIAM_FT", "=NOM_DIAM_M/0.3048")

        self.add_variable_field("Air density during tests (kg/m^3 | slug/ft^3) ",
                                "AIR_DENSITY_KG_M3", 1.225, "AIR_DENSITY_SLG_FT3",
                                "=AIR_DENSITY_KG_M3*0.00194032")

        self.add_variable_field("Air density for target descent rate (kg/m^3 | slug/ft^3)",
                                "AIR_DENSITY_TARGET_KG_M3", 1.04045,
                                "AIR_DENSITY_TARGET_SLG_FT3",
                                "=AIR_DENSITY_TARGET_KG_M3*0.00194032")

        self.add_variable_field("target descent rate (ft | m)", "DESCENT_RATE_FT_S",
                                112, "DESCENT_RATE_M_S", "=DESCENT_RATE_FT_S*0.3048")

        self.add_variable_field("Rocket Mass (lb | kg)", "ROCKET_MASS_LB",
                                100, "ROCKET_MASS_KG", "=ROCKET_MASS_LB*0.453")

        self.add_variable_field("Nominal surface area of parachute (m^2 | ft^2)",
                                "NOM_SA_M2", "=NOM_DIAM_M^2 * PI() * 0.25",
                                "NOM_SA_FT2", "=NOM_DIAM_FT^2 * PI() * 0.25")

        self.add_variable_field("Target Drag Area (m^2 | ft^2)",
                                "TARGET_DRAG_AREA_M2",
                                "=(2*ROCKET_MASS_KG)/(DESCENT_RATE_M_S^2 * AIR_DENSITY_TARGET_KG_M3)",
                                "TARGET_DRAG_AREA_FT2",
                                "=(2*ROCKET_MASS_LB)/(DESCENT_RATE_FT_S^2 * AIR_DENSITY_TARGET_SLG_FT3)")

        self.add_variable_field("Anemometer Adjustment (Calibration) Factor (unitless)",
                                "ANEMOMETER_FACTOR", 0.725)

        self.add_variable_field("Load Cell Adjustment (Calibration) Factor (unitless)",
                                "LOAD_CELL_FACTOR", 1.0)

        self.add_variable_field("Target Coeffcient of Drag Relative to Nominal SA (unitless)",
                                "TARGET_COEFF_DRAG", "=TARGET_DRAG_AREA_M2/NOM_SA_M2")

        self.add_variable_field("The windspeed threshold at which a run is said to have commenced (ft/s)",
                                "WINDSPEED_THRESHOLD", 10)

        self.add_variable_field("The force threshold at which a run is said to be commenced (lb)",
                                "FORCE_THRESHOLD", 10)

    def add_variable_field(self, title, varname_default, val_default,
                           varname_conv=None, val_conv=None):
        """
        Adds a variable field into the spreadsheet.

        Since the project is kind of a unit conversion nightmare, most values are duplicated 
        in a converted version that uses the  metric/imperial conversion. Along with adding 
        the variables into the spreadsheet using excel functions, this function also makes them
        into defined names in the entire workbook.

        All locations are determined dynamically.

        To summarize using an excel table:

        ----------------------------
        |title|val_default|val_conv|
        ----------------------------

        Where varname_default is the excel defined name for the val_default cell
        And varname_conv is the excel defined name for the val_conv cell

        Parameters
        ----------
        title: str
            The descriptor that goes into the first column
        varname_default: str
            The variable name which will be defined in excel for the default value
        val_default: str or int or float
            The value that will be in the defined cell of the default value; 
            this may be a literal or an excel formula
        varname_conv: str
            The variable name which will be defined in excel for the converted value.
            Default is None since some values do not need a conversion, in this case 
            the last column is ignored
        val_conv: str or int or float
             The value that will be in the defined cell of the converted value;
             this may be a literal or an excel formula.
             Defaul value is None for same reason as varname_conv. 
        """
        # Add raw variable input to class variables list so that there is a record
        # of all function calls and therefora all variables being managed
        self.variables_raw.append({'title':title, 'varname_default': varname_default, 
                                   'val_default':val_default, \
                                   'varname_conv':varname_conv, 'val_conv':val_conv})  
        

        row = self._next_row_idx
        self._next_row_idx += 1

        # Add the title (descriptor) of the sobject into the first column
        var_title_loc = 'A' + str(row)
        self.target_worksheet[var_title_loc] = title

        # Add the defaul
        var_default_val_location_local = 'B' + str(row)
        var_default_val_location_global = self.target_worksheet.title +\
            "!" + abs_coords(var_default_val_location_local)
        self.target_worksheet[var_default_val_location_local] = val_default
        self.target_workbook.defined_names.append(DName(varname_default,
                                                        attr_text=var_default_val_location_global))

        # For utility purposes, add the name : location into a dict
        self.defined_name_dict[varname_default] = var_default_val_location_global

        if varname_conv:
            var_conv_val_location_local = 'C' + str(row)
            self.target_worksheet[var_conv_val_location_local] = val_conv
            var_conv_val_location_global = self.target_worksheet.title +\
                "!" + abs_coords(var_conv_val_location_local)
            self.target_workbook.defined_names.append(DName(varname_conv,
                                                            attr_text=var_conv_val_location_global))

            # For utility purposes, add the name : location into a dict
            self.defined_name_dict[varname_conv] = var_conv_val_location_global
