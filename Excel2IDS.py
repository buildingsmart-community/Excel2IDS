# Excel2IDS
# Copyright (C) 2024 Artur Tomczak <artur.tomczak@buildingsmart.org>

import os
import sys
import time
import openpyxl
from ifctester import ids
from tqdm import tqdm


### settings:
SHEET_NAME = 'Requirements'
IFC_VERSION = "IFC4X3_ADD2"
ID_COL = 'A'
# applicability:
NAME_ROW = 14
ENTITY_ROW = 5
PSET_CODE_ROW = 7
PROP_CODE_ROW = 8
PVAL_CODE_ROW = 9
# req
PSET_COL = 'C'
PROP_COL = 'D'
PVAL_COL = 'E'
PVAL_PATTERN_COL = 'F'
URI_COL = 'G'
DTYPE_COL = 'H'
# split:
DISC_COL = "K"
MILESTONE = "LOI400"

ids_list = {}


def excel2ids(spreadsheet, ids_path):

    sheet = spreadsheet[SHEET_NAME]
    # Define the starting cell of the assignment table
    start_row = sheet[START_CELL].row
    start_col = sheet[START_CELL].column
    if END_CELL:
        end_row = sheet[END_CELL].row+1
        end_col = sheet[END_CELL].column+1
    else:
        end_row = sheet.max_row + 1
        end_col = sheet.max_column + 1

    for col in tqdm(range(start_col, end_col), desc="Processing Excel columns."):

        column_letter = openpyxl.utils.get_column_letter(col)
        specification_name = sheet[f"{column_letter}{NAME_ROW}"].value

        ### add applicability
        applicability = []
        entity_cell = sheet[f'{column_letter}{ENTITY_ROW}'].value
        if entity_cell:
            if '.' in entity_cell:
                entity = ids.Entity(name=entity_cell.split('.')[0].upper(), predefinedType=entity_cell.split('.')[1].upper())
            else: 
                entity = ids.Entity(name=split_multiline(entity_cell.upper()))
            applicability.append(entity)
        if sheet[f'{column_letter}{PROP_CODE_ROW}'].value:
            property = ids.Property(
                propertySet=sheet[f'{column_letter}{PSET_CODE_ROW}'].value.strip(), 
                baseName=sheet[f'{column_letter}{PROP_CODE_ROW}'].value.strip(),
                cardinality="required", # TODO should be optional
                dataType='IFCLABEL',
                )
            pv = sheet[f"{column_letter}{PVAL_CODE_ROW}"].value
            if pv:
                property.value = split_multiline(pv.strip())
                # assuming that single values (not lists) are patterns:
                if isinstance(property.value, str):
                    property.value = ids.Restriction(options={"pattern": pv.strip()})
                property.instructions = f"All objects with code like: '{pv.strip()}'"
            applicability.append(property)
        # TODO add classification
        # TODO add material

        for row in range(start_row, end_row):
            cell_value = sheet.cell(row=row, column=col).value
            ### add requirement(s)
            requirements = []
            if cell_value in ["X", "x"]:
                if sheet[f'{PROP_COL}{row}'].value:
                    property = ids.Property(
                    propertySet=sheet[f'{PSET_COL}{row}'].value.strip(), 
                    baseName=sheet[f'{PROP_COL}{row}'].value.strip(),
                    dataType=sheet[f'{DTYPE_COL}{row}'].value.strip().upper(),
                    cardinality="required"
                    )
                    if sheet[f'{PVAL_COL}{row}'].value:
                        property.value = sheet[f'{PVAL_COL}{row}'].value.strip()
                    elif sheet[f'{PVAL_PATTERN_COL}{row}'].value:
                        property.value = ids.Restriction(options={"pattern": sheet[f'{PVAL_PATTERN_COL}{row}'].value})
                    if sheet[f'{URI_COL}{row}'].value:
                        property.uri = sheet[f'{URI_COL}{row}'].value.strip()
                    requirements.append(property)
                # TODO add classifications
                # TODO add material
                # TODO add entity/predefined type
                # ids.Entity(name=sheet[f'{column_letter}{ROW_IFC}'].value.upper())
            elif cell_value:
                # assuming if a value is not 'x' it is list of possible entities
                if '.' in cell_value:
                    entity = ids.Entity(name=cell_value.split('.')[0].upper(), predefinedType=cell_value.split('.')[1].upper())
                else:
                    entity = ids.Entity(name=split_multiline(cell_value.upper()))
                requirements.append(entity)
            else:
                # skip an empty cell
                pass

            if requirements:
                disciplines = sheet[f"{DISC_COL}{row}"].value.split(",")
                for discipline in disciplines:
                    add_to_ids(
                        discipline,
                        specification_name,
                        applicability,
                        requirements,
                        purpose=discipline,
                        milestone=MILESTONE,
                    )

    ### Save all IDSes to files:
    for new_ids in tqdm(ids_list, desc="Generating separate .ids files."):
        ids_list[new_ids].to_xml(ids_path.replace(".ids", "_" + new_ids + ".ids"))
    print(
        f"\n\033[92mSuccess! {len(ids_list)} IDS files were saved in {os.path.dirname(ids_path)}.\033[0m"
    )

    # Close the spreadsheet
    spreadsheet.close()


def add_to_ids(
    ids_name,
    specification_name,
    applicability,
    requirements,
    purpose="General specification",
    milestone="Handover",
):
    """Add this specifiction to IDS file. If such IDS doesn't exist yet, create it."""

    if not ids_name in ids_list:
        # create new IDS
        ids_list[ids_name] = ids.Ids(
            title="This is an IDS experiment",
            author="technical@buildingsmart.org",
            version="0.1",
            description="This IDS was generated by a script based on the spreadsheet input.",
            date="2024-04-29",
            purpose=purpose,
            milestone=milestone,
        )

    # check if this IDS already has such category (applicability)
    exists = False
    for spec in ids_list[ids_name].specifications:
        if spec.name == specification_name:
            # ADD req!
            spec.requirements.extend(requirements)
            exists = True

    if not exists:
        # create new spec
        new_spec = ids.Specification(
            name=specification_name,
            ifcVersion=IFC_VERSION,
            # identifier=str(sheet[f'{ID_COL}{row}'].value.strip()),
        )
        new_spec.applicability = applicability
        new_spec.requirements = requirements
        ids_list[ids_name].specifications.append(new_spec)


def split_multiline(cell_value):
    # if there are multiple lines in a single cell, split it into enumeration of literal values or patterns
    if "\n" in cell_value:
        cell_value = ids.Restriction(options={"enumeration": cell_value.split("\n")})
    return cell_value


def ask_for_path():
    file_path = input(
        "\n\033[94mPlease enter the path to the Excel spreadsheet: \033[0m\n"
    )
    if file_path[0] == '"':
        file_path = file_path[1:]
    if file_path[-1] == '"':
        file_path = file_path[:-1]
    if file_path[-5:] != '.xlsx':
        print("\033[31m\nThe file must be an .xlsx. Please check the path and try again.\033[0m")
        ask_for_path()

    try:
        spreadsheet = openpyxl.load_workbook(file_path)
        return spreadsheet, file_path
    except FileNotFoundError:
        print("\033[31m\nThe file was not found. Please check the path and try again.\033[0m")
        ask_for_path()
    except Exception as e:
        print(f"\033[31m\nAn error occurred: {e}\033[0m")
        print("\n\033[94mThe program will close automatically in 10 seconds...\033[0m\n")
        time.sleep(10)
        sys.exit()


if __name__ == "__main__": 

    spreadsheet, file_path = ask_for_path()
    ids_path = file_path.replace(".xlsx", ".ids")

    START_CELL = input(
        "\n\033[94mPlease enter the name of the first cell (top-left) of the assignment table (or hit enter to use default: L15): \033[0m\n"
    )  # 'L15'
    if not START_CELL:
        START_CELL = 'L15'
    END_CELL = input(
        "\n\033[94mPlease enter the name of the last cell (bottom-right) of the assignment table (or hit enter to use default: last bottom-right cell): \033[0m\n"
    )  # 'AS54'        # or leave empty for whole table

    excel2ids(spreadsheet, ids_path)

    time.sleep(1)
    print("\n\033[94mThe program will close automatically in 10 seconds...\033[0m\n")
    time.sleep(10)
    sys.exit()