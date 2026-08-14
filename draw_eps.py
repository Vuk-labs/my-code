from pyautocad import Autocad, APoint
import win32com.client
import xlwings as xw
import time  # Import the time module for adding delays


def get_calculated_values(excel_file_path, sheet_name, x_range, y_range):
    # Open the Excel file
    wb = xw.Book(excel_file_path)

    # Access the specified sheet
    sheet = wb.sheets[sheet_name]

    # Get the calculated values for the specified ranges
    x_values = sheet.range(x_range).value
    y_values = sheet.range(y_range).value

    # Ensure that the values are always returned as lists
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]

    return x_values, y_values

def get_single_value(excel_file_path, sheet_name, cell):
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets[sheet_name]
    return sheet.range(cell).value


def insert_block(acad, block_name, insertion_point):
    try:
        # Insert the block at the specified insertion point
        acad.model.InsertBlock(APoint(*insertion_point), block_name, 1.0, 1.0, 1.0, 0)
        print(f"Block '{block_name}' inserted at {insertion_point}.")

        # Add a small delay (0.2 seconds) between commands
        time.sleep(0.2)

    except Exception as e:
        print(f"Error inserting block: {e}")


def get_block_name(excel_file_path, sheet_name, cell):
    try:
        # Open the Excel file
        wb = xw.Book(excel_file_path)

        # Access the specified sheet
        sheet = wb.sheets[sheet_name]

        # Get the block name from the specified cell
        block_name = sheet.range(cell).value

        return block_name
    except Exception as e:
        print(f"Error getting block name from Excel: {e}")
        return None


def draw_polyline(acad, layer, coordinates):
    try:
        # Set the active layer
        acad.ActiveDocument.ActiveLayer = acad.ActiveDocument.Layers(layer)

        # Format the coordinates for the pline command
        coord_str = ' '.join([f'{x},{y}' for x, y in coordinates])

        # Execute the PLine command with coordinates, specifying "c" to close the polyline
        acad.ActiveDocument.SendCommand(f'PLine {coord_str} c\n')

        print("Polyline added to AutoCAD file.")

        # Add a small delay (0.3 seconds) between commands
        time.sleep(0.3)

    except Exception as e:
        print(f"Error adding polyline to AutoCAD file: {e}")


def draw_rectangle(acad, layer, length, width, position):
    try:
        # Set the active layer
        acad.ActiveDocument.ActiveLayer = acad.ActiveDocument.Layers(layer)

        # Execute the Rectang command with position
        acad.ActiveDocument.SendCommand(f'Rectang {position[0]},{position[1]} {position[0] + length},'
                                        f'{position[1] + width}\n')
        print("Rectangle added to AutoCad file.")

        # Add a small delay (0.2 seconds) between commands
        time.sleep(0.3)

    except Exception as e:
        print(f"Error adding rectangle to AutoCAD file:{e}")


def draw_circle(acad, layer, center, radius):
    try:
        # Set the active layer
        acad.ActiveDocument.ActiveLayer = acad.ActiveDocument.Layers(layer)

        # Execute the Circle command with center and radius
        acad.ActiveDocument.SendCommand(f'Circle {center[0]},{center[1]} {radius}\n')
        print("Circle added to AutoCAD file.")

        # Add a small delay (0.2 seconds) between commands
        time.sleep(0.3)

    except Exception as e:
        print(f"Error adding circle to AutoCAD file: {e}")


def add_dimension_line(acad, start_point, end_point, at_coordinates):
    try:
        # Ensure the EX-ST-WALLS layer exists and set it as the current layer
        layers = acad.ActiveDocument.Layers
        layer_name = "EX-ST-WALLS"
        if layer_name not in [layer.Name for layer in layers]:
            layers.Add(layer_name)

        acad.ActiveDocument.ActiveLayer = layers[layer_name]

        # Execute the DimLinear command to add a linear dimension
        acad.ActiveDocument.SendCommand(f'_DimLinear {start_point[0]},{start_point[1]} {end_point[0]},{end_point[1]} '
                                        f'{at_coordinates[0]},{at_coordinates[1]}\n')

        print(f"Dimension line added to layer '{layer_name}' in the AutoCAD file.")

        # Add a small delay between commands
        time.sleep(0.4)

    except Exception as e:
        print(f"Error adding dimension line to AutoCAD file: {e}")


if __name__ == "__main__":
    # PATH TO THE EXCEL FILE CONTAINING DATA
    excel_file_path = r"C:\Users\vvvel\OneDrive\Desktop\SIGMA PM\Projects\Python\Python 00.xlsm"

    # CREATE A CONNECTION TO AUTOCAD
    acad_com = win32com.client.Dispatch("AutoCAD.Application")

    # RETRIVE THE ACTIVE EXCEL SHEET
    wb = xw.Book(excel_file_path)
    sheet = wb.sheets.active




    # Retrieve the layer names
    polyline_sheet = wb.sheets["Polylines"]
    layer_1 = polyline_sheet["A1"].value
    layer_2 = polyline_sheet["D1"].value
    layer_4 = polyline_sheet["J1"].value




    # List of circle data
    circle_data = [
        {"center_x": "V40", "center_y": "V41", "radius": "V42"},
        {"center_x": "W40", "center_y": "W41", "radius": "V42"},
        {"center_x": "X40", "center_y": "X41", "radius": "V42"},
        {"center_x": "Y40", "center_y": "Y41", "radius": "V42"},
    ]

    # Process each set of circle data
    for data in circle_data:
        # Get circle center coordinates and radius from Excel
        circle_center_x, circle_center_y = get_calculated_values(excel_file_path, "Polylines", data["center_x"],
                                                                 data["center_y"])
        radius = get_single_value(excel_file_path, "Polylines", data["radius"])

        # Define circle center and radius
        circle_center = (circle_center_x[0], circle_center_y[0])
        circle_radius = radius

        # Draw the circle
        draw_circle(acad_com, layer_2, circle_center, circle_radius)





    # Define polyline data as tuples for iteration
    polylines = [
        ("Plan Left", "B2:B9", "C2:C9"),
        ("Section Left", "B12:B19", "C12:C19"),
        ("Plan Right", "E2:E9", "F2:F9"),
        ("Section Right", "E12:E19", "F12:F19"),
        ("Section", "B22:B29", "C22:C29"),
        ("Concrete", "K2:K9", "L2:L9"),
        ("Line left", "B33:B34", "C33:C34"),
        ("Line right", "E33:E34", "F33:F34")
    ]

    # Iterate over each polyline type
    for polyline_name, x_range, y_range in polylines:
        # Get the calculated values for the ranges
        x_values, y_values = get_calculated_values(excel_file_path, "Polylines", x_range, y_range)

        # Combine x and y coordinates into tuples
        coordinates = list(zip(x_values, y_values))

        # Draw the polyline with the corresponding layer
        if polyline_name == "Concrete":
            draw_polyline(acad_com, layer_4, coordinates)
        else:
            draw_polyline(acad_com, layer_1, coordinates)





    # GET THE CALCULATED VALUES FOR THE RANGES FOR THE PLAN CENTER POLYLINE
    x_values_3, y_values_3 = get_calculated_values(excel_file_path, "Polylines", "H2:H9",
                                                   "I2:I9")

    # GET THE CALCULATED VALUES FOR THE DIMENSION LINE DATA FOR THE PLAN CENTER POLYLINE
    start_point_x_values, start_point_y_values = get_calculated_values(excel_file_path, "Polylines", "H8", "I8")
    end_point_x_values, end_point_y_values = get_calculated_values(excel_file_path, "Polylines", "H5", "I5")
    at_x_values, at_y_values = get_calculated_values(excel_file_path, "Polylines", "H8", "G2")

    # DEFINE THE COORDINATES FOR THE DIMENSION LINE FOR THE PLAN CENTER POLYLINE
    start_point_m = (start_point_x_values[0], start_point_y_values[0])
    end_point_m = (end_point_x_values[0], end_point_y_values[0])
    at_coordinates_m = (at_x_values[0], at_y_values[0])

    # COMBINE X AND Y COORDINATES INTO TUPLES PLAN CENTER POLYLINE
    coordinates_3 = list(zip(x_values_3, y_values_3))

    # NUMBER OF CENTER BLOCKS (repetitions)
    repetitions = int(polyline_sheet["D38"].value)

    # ACQUIRE THE VALUE OF INCREMENT
    horizontal_movement = polyline_sheet["G10"].value

    # REPEAT DRAWING THE PLAN CENTER POLYLINE FOR THE NUMER OF BLOCKS
    for i in range(repetitions):
        # ADJUST THE X COORDINATES FOR EACH REPETITION
        adjusted_coordinates_3b = [(x + horizontal_movement * i, y) for x, y in coordinates_3]

        # ADJUST THE DIMENSION LINE DATA FOR EACH REPETITION
        adjusted_start_point = (start_point_m[0] + horizontal_movement * i, start_point_m[1])
        adjusted_end_point = (end_point_m[0] + horizontal_movement * i, end_point_m[1])
        adjusted_at_coordinates = (at_coordinates_m[0] + horizontal_movement * i, at_coordinates_m[1])

        time.sleep(0.4)

        # DRAW POLYLINES AND DIMENSION LINES
        draw_polyline(acad_com, layer_1, adjusted_coordinates_3b)
        time.sleep(0.4)
        add_dimension_line(acad_com, adjusted_start_point, adjusted_end_point, adjusted_at_coordinates)
        time.sleep(0.4)




    # GET THE CALCULATED VALUES FOR THE RANGES FOR THE SECTION CENTER POLYLINE
    x_values_3b, y_values_3b = get_calculated_values(excel_file_path, "Polylines", "H12:H19",
                                                     "I12:I19")

    # GET THE CALCULATED VALUES FOR THE DIMENSION LINE DATA FOR THE SECTION CENTER POLYLINE
    start_point_x_values, start_point_y_values = get_calculated_values(excel_file_path, "Polylines", "H19", "I19")
    end_point_x_values, end_point_y_values = get_calculated_values(excel_file_path, "Polylines", "H14", "I14")
    at_x_values, at_y_values = get_calculated_values(excel_file_path, "Polylines", "H19", "G12")

    # DEFINE THE COORDINATES FOR THE DIMENSION LINE FOR THE PLAN CENTER POLYLINE
    start_point_mb = (start_point_x_values[0], start_point_y_values[0])
    end_point_mb = (end_point_x_values[0], end_point_y_values[0])
    at_coordinates_mb = (at_x_values[0], at_y_values[0])

    # COMBINE X AND Y COORDINATES INTO TUPLES SECTION CENTER POLYLINE
    coordinates_3 = list(zip(x_values_3b, y_values_3b))

    # REPEAT DRAWING THE SECTION CENTER POLYLINE FOR THE NUMER OF BLOCKS
    for i in range(repetitions):
        # Adjust X coordinates for each repetition
        adjusted_coordinates_3b = [(x + horizontal_movement * i, y) for x, y in coordinates_3]
        time.sleep(0.1)

        # ADJUST THE DIMENSION LINE DATA FOR EACH REPETITION
        adjusted_start_point = (start_point_mb[0] + horizontal_movement * i, start_point_mb[1])
        time.sleep(0.1)
        adjusted_end_point = (end_point_mb[0] + horizontal_movement * i, end_point_mb[1])
        time.sleep(0.1)
        adjusted_at_coordinates = (at_coordinates_mb[0] + horizontal_movement * i, at_coordinates_mb[1])

        time.sleep(0.4)

        # DRAW POLYLINES AND DIMENSION LINES
        draw_polyline(acad_com, layer_1, adjusted_coordinates_3b)
        time.sleep(0.4)
        add_dimension_line(acad_com, adjusted_start_point, adjusted_end_point, adjusted_at_coordinates)
        time.sleep(0.4)




    # LIST OF PARAMETERS FOR DIMENSION LINES
    coordinates = [
        {"start_x": "K7", "start_y": "L7", "end_x": "K6", "end_y": "L6", "at_x": "K7", "at_y": "J2",
         "name": "first_length"},
        {"start_x": "B8", "start_y": "C8", "end_x": "B5", "end_y": "C5", "at_x": "B8", "at_y": "A2",
         "name": "plan LDL"},
        {"start_x": "B19", "start_y": "C19", "end_x": "B14", "end_y": "C14", "at_x": "B19", "at_y": "A12",
         "name": "section LDL"},
        {"start_x": "E6", "start_y": "F6", "end_x": "E3", "end_y": "F3", "at_x": "E6", "at_y": "D2",
         "name": "plan RDL"},
        {"start_x": "E18", "start_y": "F18", "end_x": "E13", "end_y": "F13", "at_x": "B19", "at_y": "D12",
         "name": "section RDL"},
        {"start_x": "K9", "start_y": "L9", "end_x": "K4", "end_y": "L4", "at_x": "K9", "at_y": "J3",
         "name": "second_length"},
        {"start_x": "K7", "start_y": "L7", "end_x": "B19", "end_y": "C19", "at_x": "K7", "at_y": "J4", "name": "ewl"},
        {"start_x": "E13", "start_y": "F13", "end_x": "K6", "end_y": "L6", "at_x": "K6", "at_y": "J5", "name": "ewr"},
        {"start_x": "H2", "start_y": "I2", "end_x": "H7", "end_y": "I7", "at_x": "G3", "at_y": "G4",
         "name": "top_height"},
        {"start_x": "G13", "start_y": "I12", "end_x": "G14", "end_y": "I17", "at_x": "G15", "at_y": "I15",
         "name": "bot_height"},
        {"start_x": "G15", "start_y": "I16", "end_x": "G15", "end_y": "G16", "at_x": "G15", "at_y": "I15",
         "name": "sbot_height"},
        {"start_x": "H15", "start_y": "I15", "end_x": "H16", "end_y": "I16", "at_x": "G17", "at_y": "G18",
         "name": "tv_bot"},
        {"start_x": "H16", "start_y": "I16", "end_x": "H15", "end_y": "I15", "at_x": "G17", "at_y": "G19",
         "name": "th_bot"},
        {"start_x": "H6", "start_y": "I6", "end_x": "H5", "end_y": "I5", "at_x": "G5", "at_y": "G6", "name": "th_top"},
        {"start_x": "H7", "start_y": "I6", "end_x": "H7", "end_y": "X22", "at_x": "G3", "at_y": "R17",
         "name": "swall_dim"},
        {"start_x": "H6", "start_y": "I6", "end_x": "H5", "end_y": "I5", "at_x": "G7", "at_y": "G8", "name": "tv_top"},
        {"start_x": "H6", "start_y": "I6", "end_x": "G9", "end_y": "D9", "at_x": "D8", "at_y": "D7", "name": "top_sw"},
        {"start_x": "B9", "start_y": "C9", "end_x": "B2", "end_y": "C2", "at_x": "A3", "at_y": "A4", "name": "C1h"},
        {"start_x": "Q21", "start_y": "R21", "end_x": "Q22", "end_y": "R22", "at_x": "Q23", "at_y": "R23", "name": "C2h"},
        {"start_x": "Q3", "start_y": "R3", "end_x": "Q4", "end_y": "R4", "at_x": "Q5", "at_y": "R5", "name": "C3h"},
        {"start_x": "Q6", "start_y": "R6", "end_x": "Q7", "end_y": "R7", "at_x": "Q8", "at_y": "R8", "name": "C4h"},
        {"start_x": "Q9", "start_y": "R9", "end_x": "Q10", "end_y": "R10", "at_x": "Q11", "at_y": "R11", "name": "C1v"},
        {"start_x": "Q12", "start_y": "R12", "end_x": "Q13", "end_y": "R13", "at_x": "Q14", "at_y": "R14",
         "name": "C2v"},
        {"start_x": "Q15", "start_y": "R15", "end_x": "Q16", "end_y": "R16", "at_x": "Q17", "at_y": "R17",
         "name": "C3v"},
        {"start_x": "Q18", "start_y": "R18", "end_x": "Q19", "end_y": "R19", "at_x": "Q20", "at_y": "R20",
         "name": "C4v"},
        {"start_x": "W3", "start_y": "X3", "end_x": "W4", "end_y": "X4", "at_x": "W5", "at_y": "X5", "name": "tsmall"},
        {"start_x": "W6", "start_y": "X6", "end_x": "W7", "end_y": "X7", "at_x": "W8", "at_y": "X8", "name": "tbig"},
        {"start_x": "W9", "start_y": "X9", "end_x": "W10", "end_y": "X10", "at_x": "W11", "at_y": "X11",
         "name": "bsmall"},
        {"start_x": "W12", "start_y": "X12", "end_x": "W13", "end_y": "X13", "at_x": "W14", "at_y": "X14",
         "name": "bbig"},
        {"start_x": "W15", "start_y": "X15", "end_x": "W16", "end_y": "X16", "at_x": "W17", "at_y": "X17",
         "name": "sec_height"},
        {"start_x": "W18", "start_y": "X18", "end_x": "W19", "end_y": "X19", "at_x": "W20", "at_y": "X20",
         "name": "sec_eps"},
        {"start_x": "W21", "start_y": "X21", "end_x": "W22", "end_y": "X22", "at_x": "W23", "at_y": "X23",
         "name": "sec_width"}
    ]

    # LOOP THROUGH THE PARAMETERS AND GET CALCULATED VALUES FOR DIMENSION LINES
    for coord in coordinates:
        start_point_x_values, start_point_y_values = get_calculated_values(excel_file_path, "Polylines",
                                                                           coord["start_x"], coord["start_y"])
        end_point_x_values, end_point_y_values = get_calculated_values(excel_file_path, "Polylines", coord["end_x"],
                                                                       coord["end_y"])
        at_x_values, at_y_values = get_calculated_values(excel_file_path, "Polylines", coord["at_x"], coord["at_y"])

        # DEFINE THE COORDINATES FOR THE DIMENSION LINES
        start_point = (start_point_x_values[0], start_point_y_values[0])
        end_point = (end_point_x_values[0], end_point_y_values[0])
        at_coordinates = (at_x_values[0], at_y_values[0])

        time.sleep(0.5)

        # ADD THE DIMENSION LINES
        add_dimension_line(acad_com, start_point, end_point, at_coordinates)

        time.sleep(0.5)




    # Get the calculated values for the start point, end point, and "At" coordinates
    start_point_x_values, start_point_y_values = get_calculated_values(excel_file_path, "Polylines", "B14", "C14")
    end_point_x_values, end_point_y_values = get_calculated_values(excel_file_path, "Polylines", "H19", "I19")
    at_x_values, at_y_values = get_calculated_values(excel_file_path, "Polylines", "B14", "J6")

    # Define the start point, end point, and "At" coordinates
    start_point_mw = (start_point_x_values[0], start_point_y_values[0])
    end_point_mw = (end_point_x_values[0], end_point_y_values[0])
    at_coordinates_mw = (at_x_values[0], at_y_values[0])

    # Combine X and Y coordinates into tuples
    coordinates_3 = list(zip(x_values_3b, y_values_3b))

    # Get the number of repetitions from cell C8 in the active sheet
    repetitions = int(polyline_sheet["D38"].value) + 1

    # Get the value from cell G10 for horizontal movement
    horizontal_movement = polyline_sheet["G10"].value

    # Repeat drawing the third polyline for the specified number of times
    for i in range(repetitions):
        # Adjust X coordinates for each repetition
        adjusted_coordinates_3b = [(x + horizontal_movement * i, y) for x, y in coordinates_3]

        # Adjust start, end, and "At" coordinates for each repetition for the dimension line
        adjusted_start_point_mw = (start_point_mw[0] + horizontal_movement * i, start_point_mw[1])
        adjusted_end_point_mw = (end_point_mw[0] + horizontal_movement * i, end_point_mw[1])
        adjusted_at_coordinates_mw = (at_coordinates_mw[0] + horizontal_movement * i, at_coordinates_mw[1])

        time.sleep(0.3)

        add_dimension_line(acad_com, adjusted_start_point_mw, adjusted_end_point_mw, adjusted_at_coordinates_mw)

        time.sleep(0.3)

    time.sleep(0.3)


    # Retrieve the value from Excel cells for the rectangles
    length_specific_1 = sheet["D18"].value * 1000
    width_specific_1 = sheet["D19"].value * 1000
    length_specific_2 = sheet["D20"].value * 1000
    width_specific_2 = sheet["D19"].value * 1000

    # Define the rectangle list
    rectangles = [
        {"layer": "0", "length": 23998, "width": 14494, "position": (0, 0)},
        {"layer": "ST-SLAB EDGE", "length": length_specific_1, "width": width_specific_1, "position": (2500, 9350)},
        {"layer": "ST-SLAB EDGE", "length": length_specific_2, "width": width_specific_2,
         "position": (2500 + float(sheet["D18"].value) * 1000 + 3600, 9350)}
    ]

    # Draw rectangle in AutoCAD
    for rectangle in rectangles:
        layer = rectangle.get("layer", "ST_SLAB EDGE")
        length = rectangle["length"]
        width = rectangle["width"]
        position = rectangle["position"]
        if "layer" in rectangle:
            layer = rectangle["layer"]
        draw_rectangle(acad_com, layer, length, width, position)
        time.sleep(0.3)

    time.sleep(0.3)

    # Close the connection to AutoCAD using win32com
    acad_com = None

    time.sleep(0.2)

    # Connect to AutoCAD
    acad_pyautocad = Autocad(create_if_not_exists=True)

    # Specify the sheet name and cells containing the block name and coordinates
    sheet_name = "Polylines"

    # Define a list of dictionaries containing block name cell, x range, and y range
    blocks = [
        {"block_name_cell": "P27", "x_range": "P28", "y_range": "P29"},
        {"block_name_cell": "P30", "x_range": "P31", "y_range": "P32"},
        {"block_name_cell": "P33", "x_range": "P34", "y_range": "P35"},
        {"block_name_cell": "P36", "x_range": "P37", "y_range": "P38"}
    ]

    # Loop through each block definition
    for block in blocks:
        block_name_cell = block["block_name_cell"]
        x_range = block["x_range"]
        y_range = block["y_range"]

        # Retrieve the block name from Excel
        block_name = get_block_name(excel_file_path, sheet_name, block_name_cell)

        # Retrieve the coordinates from Excel
        x_coordinates, y_coordinates = get_calculated_values(excel_file_path, sheet_name, x_range, y_range)

        if block_name and x_coordinates and y_coordinates:
            # Specify the insertion point as the first pair of coordinates
            insertion_point = (x_coordinates[0], y_coordinates[0], 0)

            # Insert the block into the drawing
            insert_block(acad_pyautocad, block_name, insertion_point)
        else:
            print("Failed to retrieve block name or coordinates from Excel.")

    # Retrieve pontoon width
    pontoon_width = sheet["D19"].value

    if pontoon_width >= 4:
        # Define block information including sheet name, block name cell, and coordinate ranges
        block_info_list = [
            {"sheet_name": "Polylines", "block_name_cell": "P46", "x_range": "P47", "y_range": "P48"},
            {"sheet_name": "Polylines", "block_name_cell": "P46", "x_range": "P49", "y_range": "P50"},
            {"sheet_name": "Polylines", "block_name_cell": "P51", "x_range": "P52", "y_range": "P53"},
            {"sheet_name": "Polylines", "block_name_cell": "P51", "x_range": "P54", "y_range": "P55"},
            {"sheet_name": "Polylines", "block_name_cell": "T30", "x_range": "T31", "y_range": "T32"},
            {"sheet_name": "Polylines", "block_name_cell": "T36", "x_range": "T37", "y_range": "T38"}
        ]

        # Loop through each block information
        for block_info in block_info_list:
            # Retrieve block information
            sheet_name = block_info["sheet_name"]
            block_name_cell = block_info["block_name_cell"]
            x_range = block_info["x_range"]
            y_range = block_info["y_range"]

            # Retrieve block name and coordinates from Excel
            block_name = get_block_name(excel_file_path, sheet_name, block_name_cell)
            x_coordinates, y_coordinates = get_calculated_values(excel_file_path, sheet_name, x_range, y_range)

            if block_name and x_coordinates and y_coordinates:
                # Specify the insertion point as the first pair of coordinates
                insertion_point = (x_coordinates[0], y_coordinates[0], 0)

                # Insert the block into the drawing
                insert_block(acad_pyautocad, block_name, insertion_point)
            else:
                print("Failed to retrieve block name or coordinates from Excel.")
    else:
        # Specify the sheet name
        sheet_name = "Polylines"

        # Define block information including block name cell and coordinate ranges
        block_info_list = [
            {"block_name_cell": "P40", "x_range": "P41", "y_range": "P42"},
            {"block_name_cell": "P43", "x_range": "P44", "y_range": "P45"},
            {"block_name_cell": "T27", "x_range": "T28", "y_range": "T29"},
            {"block_name_cell": "T33", "x_range": "T34", "y_range": "T35"}
        ]

        # Loop through each block information
        for block_info in block_info_list:
            # Retrieve block information
            block_name_cell = block_info["block_name_cell"]
            x_range = block_info["x_range"]
            y_range = block_info["y_range"]

            # Retrieve block name and coordinates from Excel
            block_name = get_block_name(excel_file_path, sheet_name, block_name_cell)
            x_coordinates, y_coordinates = get_calculated_values(excel_file_path, sheet_name, x_range, y_range)

            if block_name and x_coordinates and y_coordinates:
                # Specify the insertion point as the first pair of coordinates
                insertion_point = (x_coordinates[0], y_coordinates[0], 0)

                # Insert the block into the drawing
                insert_block(acad_pyautocad, block_name, insertion_point)
            else:
                print("Failed to retrieve block name or coordinates from Excel.")

    # Specify the sheet name and cells containing the block name and coordinates
    sheet_name = "Polylines"

    # Define block information including block name cell and coordinate ranges
    block_info_list = [
        {"block_name_cell": "T53", "x_range": "T54", "y_range": "T55"},
        {"block_name_cell": "T56", "x_range": "T57", "y_range": "T58"},
        {"block_name_cell": "T60", "x_range": "T61", "y_range": "T62"},
        {"block_name_cell": "T46", "x_range": "T47", "y_range": "T48"},
        {"block_name_cell": "T49", "x_range": "T50", "y_range": "T51"},
        {"block_name_cell": "T40", "x_range": "T41", "y_range": "T42"},
        {"block_name_cell": "T43", "x_range": "T44", "y_range": "T45"},
        {"block_name_cell": "U46", "x_range": "V47", "y_range": "V48"},
        {"block_name_cell": "U49", "x_range": "V50", "y_range": "V51"},
        {"block_name_cell": "U52", "x_range": "V53", "y_range": "V54"},
        {"block_name_cell": "U55", "x_range": "V56", "y_range": "V57"}
    ]

    # Loop through each block information
    for block_info in block_info_list:
        # Retrieve block information
        block_name_cell = block_info["block_name_cell"]
        x_range = block_info["x_range"]
        y_range = block_info["y_range"]

        # Retrieve block name and coordinates from Excel
        block_name = get_block_name(excel_file_path, sheet_name, block_name_cell)
        x_coordinates, y_coordinates = get_calculated_values(excel_file_path, sheet_name, x_range, y_range)

        if block_name and x_coordinates and y_coordinates:
            # Specify the insertion point as the first pair of coordinates
            insertion_point = (x_coordinates[0], y_coordinates[0], 0)

            # Insert the block into the drawing
            insert_block(acad_pyautocad, block_name, insertion_point)
        else:
            print("Failed to retrieve block name or coordinates from Excel.")