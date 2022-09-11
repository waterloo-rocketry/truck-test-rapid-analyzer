from truck_test_rapid_analyzer import HeaderManager
from truck_test_rapid_analyzer import create_averageifs_formula


def test_averageifs():
    # Since the real formula uses quotations, they must be escaped
    correct_result = "=AVERAGEIFS(SheetName!J2:J15000,SheetName!G2:G15000, " \
        + "\">=\"&WINDSPEED_THRESHOLD, SheetName!F2:F15000, \">=\"&FORCE_THRESHOLD)"
    result = create_averageifs_formula("SheetName", 15000)
    print(correct_result)
    print(result)

    assert correct_result == result
