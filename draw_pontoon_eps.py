import datetime
import time

import pythoncom
import win32com.client
import xlwings as xw


PARKING_MARGIN = 3000

SIZE_CONFIG = {
    "small": {
        "rectangle_size": (18719, 15523),
        "rect_origin": (0, -1965),
        "concrete_start": (1500, 5000),
        "celtscale": 5,
        "dimscale": 80,
        "block_scale": 1,
        "lifting_anchor_block_y": 3500,
    },
    "medium": {
        "rectangle_size": (23398, 19404),
        "rect_origin": (0, -2455),
        "concrete_start": (1500, 7000),
        "celtscale": 12.5,
        "dimscale": 100,
        "block_scale": 1.25,
        "lifting_anchor_block_y": 4000,
    },
    "large": {
        "rectangle_size": (28078, 23285),
        "rect_origin": (0, -2945),
        "concrete_start": (1500, 9000),
        "celtscale": 25,
        "dimscale": 120,
        "block_scale": 1.5,
        "lifting_anchor_block_y": 4500,
    },
}

MAX_RECT_WIDTH = max(cfg["rectangle_size"][0] for cfg in SIZE_CONFIG.values())

STYRO_POLYLINES = [
    ("Plan Left", "B2:B9", "C2:C9"),
    ("Section Left", "B12:B19", "C12:C19"),
    ("Plan Right", "E2:E9", "F2:F9"),
    ("Section Right", "E12:E19", "F12:F19"),
    ("Section", "B22:B29", "C22:C29"),
    ("Line left", "B33:B34", "C33:C34"),
    ("Line right", "E33:E34", "F33:F34"),
]

STYRO_CENTER_POLYLINES = [
    ("Plan Center", "H2:H9", "I2:I9", ("H8", "I8"), ("H5", "I5"), ("H8", "G54")),
    ("Section Center", "H12:H19", "I12:I19", ("H19", "I19"), ("H14", "I14"), ("H19", "G50")),
]

DIMENSION_LINES = [
    {"start_x": "K7", "start_y": "L7", "end_x": "K6", "end_y": "L6", "at_x": "K7", "at_y": "G43", "name": "first_length"},
    {"start_x": "B8", "start_y": "C8", "end_x": "B5", "end_y": "C5", "at_x": "B8", "at_y": "G62", "name": "plan LDL"},
    {"start_x": "B19", "start_y": "C19", "end_x": "B14", "end_y": "C14", "at_x": "B19", "at_y": "G48", "name": "section LDL"},
    {"start_x": "E6", "start_y": "F6", "end_x": "E3", "end_y": "F3", "at_x": "E6", "at_y": "G61", "name": "plan RDL"},
    {"start_x": "E18", "start_y": "F18", "end_x": "E13", "end_y": "F13", "at_x": "B19", "at_y": "G49", "name": "section RDL"},
    {"start_x": "K9", "start_y": "L9", "end_x": "K4", "end_y": "L4", "at_x": "K9", "at_y": "G44", "name": "second_length"},
    {"start_x": "K7", "start_y": "L7", "end_x": "B19", "end_y": "C19", "at_x": "K7", "at_y": "G45", "name": "ewl"},
    {"start_x": "E13", "start_y": "F13", "end_x": "K6", "end_y": "L6", "at_x": "K6", "at_y": "G46", "name": "ewr"},
    {"start_x": "B3", "start_y": "C3", "end_x": "B6", "end_y": "C6", "at_x": "G55", "at_y": "G56", "name": "top_height"},
    {"start_x": "J43", "start_y": "K43", "end_x": "J44", "end_y": "K44", "at_x": "J45", "at_y": "I15", "name": "bot_height"},
    {"start_x": "J46", "start_y": "K46", "end_x": "J47", "end_y": "K47", "at_x": "J48", "at_y": "K48", "name": "deck_thickness"},
    {"start_x": "H15", "start_y": "I15", "end_x": "H16", "end_y": "I16", "at_x": "G17", "at_y": "G18", "name": "tv_bot"},
    {"start_x": "H16", "start_y": "I16", "end_x": "H15", "end_y": "I15", "at_x": "G17", "at_y": "G19", "name": "th_bot"},
    {"start_x": "B6", "start_y": "C6", "end_x": "B5", "end_y": "C5", "at_x": "G57", "at_y": "G58", "name": "th_top"},
    {"start_x": "B6", "start_y": "C6", "end_x": "B6", "end_y": "X22", "at_x": "G55", "at_y": "R17", "name": "swall_dim"},
    {"start_x": "B6", "start_y": "C6", "end_x": "B5", "end_y": "C5", "at_x": "G59", "at_y": "G60", "name": "tv_top"},
    {"start_x": "Q24", "start_y": "R24", "end_x": "Q25", "end_y": "R25", "at_x": "Q26", "at_y": "R26", "name": "C1h"},
    {"start_x": "Q21", "start_y": "R21", "end_x": "Q22", "end_y": "R22", "at_x": "Q23", "at_y": "R23", "name": "C2h"},
    {"start_x": "Q3", "start_y": "R3", "end_x": "Q4", "end_y": "R4", "at_x": "Q5", "at_y": "R5", "name": "C3h"},
    {"start_x": "Q6", "start_y": "R6", "end_x": "Q7", "end_y": "R7", "at_x": "Q8", "at_y": "R8", "name": "C4h"},
    {"start_x": "Q9", "start_y": "R9", "end_x": "Q10", "end_y": "R10", "at_x": "Q11", "at_y": "R11", "name": "C1v"},
    {"start_x": "Q12", "start_y": "R12", "end_x": "Q13", "end_y": "R13", "at_x": "Q14", "at_y": "R14", "name": "C2v"},
    {"start_x": "Q15", "start_y": "R15", "end_x": "Q16", "end_y": "R16", "at_x": "Q17", "at_y": "R17", "name": "C3v"},
    {"start_x": "Q18", "start_y": "R18", "end_x": "Q19", "end_y": "R19", "at_x": "Q20", "at_y": "R20", "name": "C4v"},
    {"start_x": "W3", "start_y": "X3", "end_x": "W4", "end_y": "X4", "at_x": "W5", "at_y": "X5", "name": "tsmall"},
    {"start_x": "W6", "start_y": "X6", "end_x": "W7", "end_y": "X7", "at_x": "W8", "at_y": "X8", "name": "tbig"},
    {"start_x": "W9", "start_y": "X9", "end_x": "W10", "end_y": "X10", "at_x": "W11", "at_y": "X11", "name": "bsmall"},
    {"start_x": "W12", "start_y": "X12", "end_x": "W13", "end_y": "X13", "at_x": "W14", "at_y": "X14", "name": "bbig"},
    {"start_x": "W15", "start_y": "X15", "end_x": "W16", "end_y": "X16", "at_x": "W17", "at_y": "X17", "name": "sec_height"},
    {"start_x": "W18", "start_y": "X18", "end_x": "W19", "end_y": "X19", "at_x": "W20", "at_y": "X20", "name": "sec_eps"},
    {"start_x": "W21", "start_y": "X21", "end_x": "W22", "end_y": "X22", "at_x": "W23", "at_y": "X23", "name": "sec_width"},
]


def get_circle_coordinate_range(circle_sum):
    if circle_sum % 2 == 0 and circle_sum != 6:
        return "AA41:AA46", "AB41:AB46"
    else:
        return "AD41:AD48", "AE41:AE48"


def get_size_category(length):
    if length <= 10:
        return "small"
    elif length <= 15:
        return "medium"
    else:
        return "large"


def draw_rectangle(acad, layer, length, width, position):
    try:
        acad.ActiveDocument.ActiveLayer = acad.ActiveDocument.Layers(layer)
        acad.ActiveDocument.SendCommand(
            f'Rectang {position[0]},{position[1]} {position[0] + length},{position[1] + width}\n'
        )
        print(f"Rectangle {length}x{width} added on layer '{layer}' at {position}.")
        time.sleep(0.5)

    except Exception as e:
        print(f"Error adding rectangle to AutoCAD file: {e}")


def draw_polyline(acad, layer, coordinates):
    try:
        acad.ActiveDocument.ActiveLayer = acad.ActiveDocument.Layers(layer)
        coord_str = ' '.join(f'{x},{y}' for x, y in coordinates)
        acad.ActiveDocument.SendCommand(f'PLine {coord_str} c\n')
        print(f"Polyline added on layer '{layer}'.")
        time.sleep(0.5)

    except Exception as e:
        print(f"Error adding polyline to AutoCAD file: {e}")


def draw_circle(acad, layer, center, radius):
    try:
        acad.ActiveDocument.ActiveLayer = acad.ActiveDocument.Layers(layer)
        acad.ActiveDocument.SendCommand(f'Circle {center[0]},{center[1]} {radius}\n')
        print(f"Circle added on layer '{layer}' at {center}, radius {radius}.")
        time.sleep(0.5)

    except Exception as e:
        print(f"Error adding circle to AutoCAD file: {e}")


def get_coordinates(polylines_sheet, x_range, y_range):
    x_values = polylines_sheet[x_range].value
    y_values = polylines_sheet[y_range].value
    if not isinstance(x_values, list):
        x_values = [x_values]
    if not isinstance(y_values, list):
        y_values = [y_values]
    return list(zip(x_values, y_values))


def get_point(polylines_sheet, x_cell, y_cell):
    return polylines_sheet[x_cell].value, polylines_sheet[y_cell].value


def draw_dimension_line(acad, start_point, end_point, at_point, layer="Dimensions"):
    if None in start_point or None in end_point or None in at_point:
        print(f"Skipping dimension line on layer '{layer}': blank coordinate cell(s).")
        return

    try:
        layers = acad.ActiveDocument.Layers
        if layer not in [l.Name for l in layers]:
            layers.Add(layer)
        acad.ActiveDocument.ActiveLayer = layers(layer)

        acad.ActiveDocument.SendCommand(
            f'_DimLinear {start_point[0]},{start_point[1]} {end_point[0]},{end_point[1]} '
            f'{at_point[0]},{at_point[1]}\n'
        )
        print(f"Dimension line added on layer '{layer}'.")
        time.sleep(1.2)

    except Exception as e:
        print(f"Error adding dimension line to AutoCAD file: {e}")


def to_variant_point(point):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [point[0], point[1], 0])


def insert_block(acad, block_name, insertion_point, layer, x_scale=1.0, y_scale=1.0, z_scale=1.0, rotation=0.0):
    try:
        layers = acad.ActiveDocument.Layers
        if layer not in [l.Name for l in layers]:
            layers.Add(layer)
        acad.ActiveDocument.ActiveLayer = layers(layer)

        block_ref = acad.ActiveDocument.ModelSpace.InsertBlock(
            to_variant_point(insertion_point), block_name, x_scale, y_scale, z_scale, rotation
        )
        print(f"Block '{block_name}' inserted on layer '{layer}' at {insertion_point}.")
        time.sleep(0.5)
        return block_ref

    except Exception as e:
        print(f"Error inserting block '{block_name}': {e}")
        return None


def mirror_entity(entity, point1, point2):
    try:
        mirrored = entity.Mirror(to_variant_point(point1), to_variant_point(point2))
        print("Entity mirrored.")
        time.sleep(0.5)
        return mirrored

    except Exception as e:
        print(f"Error mirroring entity: {e}")
        return None


def zoom_viewport_to_window(acad, layout_name, corner1, corner2, max_attempts=4, retry_delay=1.0):
    for attempt in range(1, max_attempts + 1):
        try:
            acad.ActiveDocument.ActiveLayout = acad.ActiveDocument.Layouts.Item(layout_name)
            acad.ActiveDocument.SendCommand("_MSPACE\n")
            time.sleep(0.5)
            acad.ActiveDocument.SendCommand(f"_ZOOM _W {corner1[0]},{corner1[1]} {corner2[0]},{corner2[1]}\n")
            time.sleep(0.5)
            acad.ActiveDocument.SendCommand("_PSPACE\n")
            time.sleep(0.5)
            print(f"Viewport on layout '{layout_name}' zoomed to {corner1}-{corner2}.")
            return

        except Exception as e:
            if attempt == max_attempts:
                print(f"Error zooming viewport after {max_attempts} attempts: {e}")
            else:
                print(f"Zoom attempt {attempt} failed ({e}), retrying...")
                time.sleep(retry_delay)


def archive_pontoon(acad, layout_name, entities, corner1, corner2, margin):
    try:
        layout_names = [l.Name for l in acad.ActiveDocument.Layouts]
        other_pontoons = [
            name for name in layout_names if name not in ("Model", "Original layout", layout_name)
        ]
        slot = len(other_pontoons)

        offset = (slot * (MAX_RECT_WIDTH + margin), 0)

        for entity in entities:
            entity.Move(to_variant_point((0, 0)), to_variant_point(offset))

        new_corner1 = (corner1[0] + offset[0], corner1[1] + offset[1])
        new_corner2 = (corner2[0] + offset[0], corner2[1] + offset[1])
        zoom_viewport_to_window(acad, layout_name, new_corner1, new_corner2)

        layout_block = acad.ActiveDocument.Layouts.Item(layout_name).Block
        for entity in layout_block:
            if entity.ObjectName == "AcDbViewport":
                entity.DisplayLocked = True

        print(f"Archived {len(entities)} entities on layout '{layout_name}' to slot {slot}, shifted by {offset}.")

    except Exception as e:
        print(f"Error archiving pontoon: {e}")


def update_date_text(acad, layout_name, prefix="Date:"):
    try:
        today_str = datetime.datetime.now().strftime("%d/%m/%y")
        layout_block = acad.ActiveDocument.Layouts.Item(layout_name).Block

        found_texts = []
        for entity in layout_block:
            if entity.ObjectName in ("AcDbText", "AcDbMText"):
                found_texts.append(entity.TextString)
                if "Date" in entity.TextString:
                    entity.TextString = f"{prefix} {today_str}"
                    print(f"Updated date text to '{prefix} {today_str}'.")
                    return

        print(f"No text containing 'Date' found on layout '{layout_name}'. Text entities found: {found_texts}")

    except Exception as e:
        print(f"Error updating date text: {e}")


def replace_text_substring(acad, layout_name, search_text, new_text):
    try:
        layout_block = acad.ActiveDocument.Layouts.Item(layout_name).Block

        found_texts = []
        for entity in layout_block:
            if entity.ObjectName in ("AcDbText", "AcDbMText"):
                found_texts.append(entity.TextString)
                if search_text in entity.TextString:
                    entity.TextString = entity.TextString.replace(search_text, str(new_text))
                    print(f"Replaced '{search_text}' substring with '{new_text}'.")
                    return

        print(f"No text containing '{search_text}' found on layout '{layout_name}'. Text entities found: {found_texts}")

    except Exception as e:
        print(f"Error replacing text substring '{search_text}': {e}")


def replace_layout_text(acad, layout_name, search_text, new_text):
    try:
        layout_block = acad.ActiveDocument.Layouts.Item(layout_name).Block

        found_texts = []
        for entity in layout_block:
            if entity.ObjectName in ("AcDbText", "AcDbMText"):
                found_texts.append(entity.TextString)
                stripped = entity.TextString.strip()
                if search_text.lower() in stripped.lower() and not stripped.endswith(":"):
                    entity.TextString = str(new_text)
                    print(f"Replaced '{search_text}' text with '{new_text}'.")
                    return

        print(f"No non-label text containing '{search_text}' found on layout '{layout_name}'. Text entities found: {found_texts}")

    except Exception as e:
        print(f"Error replacing layout text '{search_text}': {e}")


if __name__ == "__main__":
    acad_com = win32com.client.Dispatch("AutoCAD.Application")
    acad_com.ActiveDocument.SetVariable("PSLTSCALE", 0)
    acad_com.ActiveDocument.SetVariable("TILEMODE", 1)
    original_osmode = acad_com.ActiveDocument.GetVariable("OSMODE")
    acad_com.ActiveDocument.SetVariable("OSMODE", 0)
    time.sleep(0.5)

    initial_handles = {e.Handle for e in acad_com.ActiveDocument.ModelSpace}

    sheet = xw.sheets.active

    length = sheet["D18"].value
    width = sheet["D19"].value
    length_2 = sheet["D20"].value

    config = SIZE_CONFIG[get_size_category(length)]

    rect_length, rect_width = config["rectangle_size"]
    rect_origin = config["rect_origin"]
    draw_rectangle(acad_com, "0", rect_length, rect_width, rect_origin)

    concrete_length = length * 1000
    concrete_width = width * 1000
    concrete_length_2 = length_2 * 1000

    x1, y1 = config["concrete_start"]
    x2 = 5000 + length * 1000
    y2 = y1

    draw_rectangle(acad_com, "Concrete", concrete_length, concrete_width, (x1, y1))
    draw_rectangle(acad_com, "Concrete", concrete_length_2, concrete_width, (x2, y2))

    polylines_sheet = sheet.book.sheets["Polylines"]
    x_values = polylines_sheet["K2:K9"].value
    y_values = polylines_sheet["L2:L9"].value
    coordinates = list(zip(x_values, y_values))

    draw_polyline(acad_com, "Concrete", coordinates)

    acad_com.ActiveDocument.SetVariable("CELTSCALE", config["celtscale"])

    for name, x_range, y_range in STYRO_POLYLINES:
        draw_polyline(acad_com, "Styro", get_coordinates(polylines_sheet, x_range, y_range))

    repetitions = int(polylines_sheet["D38"].value)
    horizontal_movement = polylines_sheet["D39"].value

    for name, x_range, y_range, _, _, _ in STYRO_CENTER_POLYLINES:
        base_coordinates = get_coordinates(polylines_sheet, x_range, y_range)
        for i in range(repetitions):
            shifted_coordinates = [(x + horizontal_movement * i, y) for x, y in base_coordinates]
            draw_polyline(acad_com, "Styro", shifted_coordinates)

    acad_com.ActiveDocument.SetVariable("CELTSCALE", 1)

    circle_sum = sum(value or 0 for value in sheet["C6:C12"].value)
    circle_x_range, circle_y_range = get_circle_coordinate_range(circle_sum)
    circle_radius = polylines_sheet["AA48"].value
    circle_centers = get_coordinates(polylines_sheet, circle_x_range, circle_y_range)

    for circle_center in circle_centers:
        draw_circle(acad_com, "Lifting Anchors", circle_center, circle_radius)

    lifting_anchor_block_name = str(polylines_sheet["AB48"].value).strip()
    lifting_anchor_block_y = config["lifting_anchor_block_y"]

    for pair_start in range(0, len(circle_centers), 2):
        pair_x = circle_centers[pair_start][0]
        insert_block(acad_com, lifting_anchor_block_name, (pair_x, lifting_anchor_block_y), "Lifting Anchors")

    acad_com.ActiveDocument.SetVariable("DIMSCALE", config["dimscale"])

    for name, x_range, y_range, dim_start_cells, dim_end_cells, dim_at_cells in STYRO_CENTER_POLYLINES:
        dim_start = get_point(polylines_sheet, *dim_start_cells)
        dim_end = get_point(polylines_sheet, *dim_end_cells)
        dim_at = get_point(polylines_sheet, *dim_at_cells)
        for i in range(repetitions):
            shifted_dim_start = (dim_start[0] + horizontal_movement * i, dim_start[1])
            shifted_dim_end = (dim_end[0] + horizontal_movement * i, dim_end[1])
            shifted_dim_at = (dim_at[0] + horizontal_movement * i, dim_at[1])
            draw_dimension_line(acad_com, shifted_dim_start, shifted_dim_end, shifted_dim_at)

    for dim in DIMENSION_LINES:
        start_point = get_point(polylines_sheet, dim["start_x"], dim["start_y"])
        end_point = get_point(polylines_sheet, dim["end_x"], dim["end_y"])
        at_point = get_point(polylines_sheet, dim["at_x"], dim["at_y"])
        draw_dimension_line(acad_com, start_point, end_point, at_point)

    mw_start = get_point(polylines_sheet, "B14", "C14")
    mw_end = get_point(polylines_sheet, "H19", "I19")
    mw_at = get_point(polylines_sheet, "B14", "G47")

    for i in range(repetitions + 1):
        shifted_mw_start = (mw_start[0] + horizontal_movement * i, mw_start[1])
        shifted_mw_end = (mw_end[0] + horizontal_movement * i, mw_end[1])
        shifted_mw_at = (mw_at[0] + horizontal_movement * i, mw_at[1])
        draw_dimension_line(acad_com, shifted_mw_start, shifted_mw_end, shifted_mw_at)

    acad_com.ActiveDocument.SetVariable("DIMSCALE", 1)

    block_name = str(polylines_sheet["P28"].value).strip()
    block_insertion_point = get_point(polylines_sheet, "P29", "P30")
    axis1_point1 = get_point(polylines_sheet, "P32", "Q32")
    axis1_point2 = get_point(polylines_sheet, "P33", "Q33")
    axis2_point1 = get_point(polylines_sheet, "P34", "Q34")
    axis2_point2 = get_point(polylines_sheet, "P35", "Q35")

    mirror_mode_value = sheet["Q4"].value
    if isinstance(mirror_mode_value, float) and mirror_mode_value.is_integer():
        mirror_mode = str(int(mirror_mode_value))
    else:
        mirror_mode = str(mirror_mode_value).strip().lower()

    block_0 = insert_block(acad_com, block_name, block_insertion_point, "Hardware")
    block_1 = mirror_entity(block_0, axis1_point1, axis1_point2)

    if mirror_mode in ("4", "2r"):
        block_2 = mirror_entity(block_0, axis2_point1, axis2_point2)
        block_3 = mirror_entity(block_1, axis2_point1, axis2_point2)

        if mirror_mode == "2r" and block_2 and block_3:
            block_0.Delete()
            block_1.Delete()

    second_block_name = str(polylines_sheet["P40"].value).strip()

    if width >= 4:
        second_block_insertion_points = [
            get_point(polylines_sheet, "O45", "P45"),
            get_point(polylines_sheet, "O46", "P46"),
        ]
    else:
        second_block_insertion_points = [get_point(polylines_sheet, "O42", "P42")]

    second_blocks = [
        insert_block(acad_com, second_block_name, point, "Hardware")
        for point in second_block_insertion_points
    ]

    if mirror_mode in ("4", "2r"):
        mirrored_second_blocks = [
            mirror_entity(block, axis2_point1, axis2_point2) for block in second_blocks
        ]

        if mirror_mode == "2r" and all(mirrored_second_blocks):
            for block in second_blocks:
                block.Delete()

    third_hardware_block_name = str(polylines_sheet["P36"].value).strip()
    third_hardware_block_insertion_point = get_point(polylines_sheet, "P37", "P38")
    third_hardware_block = insert_block(
        acad_com, third_hardware_block_name, third_hardware_block_insertion_point, "Hardware"
    )

    if mirror_mode in ("4", "2r"):
        mirrored_third_hardware_block = mirror_entity(third_hardware_block, axis2_point1, axis2_point2)

        if mirror_mode == "2r" and mirrored_third_hardware_block:
            third_hardware_block.Delete()

    scale = config["block_scale"]

    third_block_name = str(polylines_sheet["S54"].value).strip()
    third_block_insertion_point = get_point(polylines_sheet, "T54", "U54")
    insert_block(acad_com, third_block_name, third_block_insertion_point, "Symbols", scale, scale)

    fourth_block_name = str(polylines_sheet["S55"].value).strip()
    fourth_block_insertion_point = get_point(polylines_sheet, "T55", "U55")
    insert_block(acad_com, fourth_block_name, fourth_block_insertion_point, "Symbols", scale, scale)

    fifth_block_name = str(polylines_sheet["S56"].value).strip()
    fifth_block_insertion_point = get_point(polylines_sheet, "T56", "U56")
    insert_block(acad_com, fifth_block_name, fifth_block_insertion_point, "Symbols", scale, scale)

    sixth_block_name = str(polylines_sheet["S57"].value).strip()
    sixth_block_insertion_point = get_point(polylines_sheet, "T57", "U57")
    insert_block(acad_com, sixth_block_name, sixth_block_insertion_point, "Symbols", scale, scale)

    seventh_block_name = str(polylines_sheet["S58"].value).strip()
    seventh_block_insertion_point = get_point(polylines_sheet, "T58", "U58")
    insert_block(acad_com, seventh_block_name, seventh_block_insertion_point, "Symbols", scale, scale)

    acad_com.ActiveDocument.SendCommand("_REGEN\n")
    time.sleep(0.5)

    acad_com.ActiveDocument.SendCommand("-LAYOUT\nC\nOriginal layout\nLayout1\n")
    time.sleep(1.5)

    rect_corner2 = (rect_origin[0] + rect_length, rect_origin[1] + rect_width)
    zoom_viewport_to_window(acad_com, "Layout1", rect_origin, rect_corner2)

    update_date_text(acad_com, "Layout1")

    replace_layout_text(acad_com, "Layout1", "project name", sheet["S14"].value)
    replace_layout_text(acad_com, "Layout1", "tittle", sheet["S15"].value)

    replace_text_substring(
        acad_com,
        "Layout1",
        "TRIM EACH EPS BLOCK: TOP 180x90(HxV) AROUND, BOT. 90x180(HxV) SIDEWAYS ONLY, 200x200 VERT. EACH BLOCK CORNER UNLESS OTHERWISE NOTED",
        sheet["S18"].value,
    )
    replace_text_substring(acad_com, "Layout1", "6xRd30x450", sheet["S19"].value)
    replace_text_substring(acad_com, "Layout1", "40kN", sheet["S20"].value)
    replace_text_substring(acad_com, "Layout1", "00", sheet["S17"].value)

    new_entities = [e for e in acad_com.ActiveDocument.ModelSpace if e.Handle not in initial_handles]
    archive_pontoon(acad_com, "Layout1", new_entities, rect_origin, rect_corner2, PARKING_MARGIN)

    acad_com.ActiveDocument.SendCommand("_ZOOM _E\n")
    time.sleep(0.5)

    acad_com.ActiveDocument.Layouts.Item("Layout1").Name = str(sheet["S15"].value)

    acad_com.ActiveDocument.SetVariable("OSMODE", original_osmode)
