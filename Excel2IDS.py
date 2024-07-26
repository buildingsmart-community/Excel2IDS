# Excel2IDS
# Copyright (C) 2024 Artur Tomczak <artomczak@gmail.com>

import os
import sys
import time
import copy
import openpyxl
from ifctester import ids
from tqdm import tqdm
import json
import re


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



def not_empty(v):
    return not (v is None or v == '')


def excel2ids(spreadsheet, ids_path):
    sheet = spreadsheet[s.SHEET_NAME]
    # Define the starting cell of the assignment table
    start_row = s.DEFAULT_START_CELL.row
    start_col = s.DEFAULT_START_CELL.column
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
                        property.value = split_multiline(sheet[f'{PVAL_PATTERN_COL}{row}'].value.strip())
                        # TODO TEMP if isinstance(property.value, str):
                        # TODO TEMP property.value = ids.Restriction(options={"pattern": sheet[f'{PVAL_PATTERN_COL}{row}'].value})
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
    """Add this specifiction to IDS file. If such IDS doesn't exist yet, create it.
    Known limitations: 
    - the applicability is automatically set to 'minOccur'=0, meaning 'if exists'/'may occur' and does not trigger an error if no such element is found.
    """

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
            spec.requirements += requirements
            exists = True

    if not exists:
        # create new spec
        new_spec = ids.Specification(
            name=specification_name,
            ifcVersion=IFC_VERSION,
            # identifier=str(sheet[f'{ID_COL}{row}'].value.strip()),
        )
        new_spec.applicability = copy.deepcopy(applicability)
        new_spec.requirements = copy.deepcopy(requirements)
        ids_list[ids_name].specifications.append(new_spec)


QUOTED_PATTERN = r'^".*"$'

def process_value(cell_value):
    """ Look at the value and convert to enumeration of pattern if needed """
    # if there are multiple lines in a single cell, split it into enumeration of literal values
    if "\n" in cell_value:
        cell_value = ids.Restriction(options={"enumeration": cell_value.split("\n")})
    # if it's in quotation, turn into pattern restriction (regex):
    elif bool(re.match(QUOTED_PATTERN, cell_value)):
        cell_value = ids.Restriction(options={"pattern": cell_value[1:-1]})
    # in other cases, return as is (empty or literal value)
    return cell_value


def color_text(text, color='blue'):
    if color == 'blue':
        text = '\033[94m' + text + '\033[0m'
    elif color == 'red':
        text = '\033[31m' + text + '\033[0m'
    return text


def ask_for_path():
    file_path = input(color_text("\nPlease enter the path to the Excel spreadsheet: \n"))
    if file_path[0] == '"':
        file_path = file_path[1:]
    if file_path[-1] == '"':
        file_path = file_path[:-1]
    if file_path[-5:] != '.xlsx':
        print(color_text("\nThe file must be an .xlsx. Please check the path and try again.", color='red'))
        ask_for_path()
    try:
        spreadsheet = openpyxl.load_workbook(file_path)
        return spreadsheet, file_path
    except FileNotFoundError:
        print(color_text("\nThe file was not found. Please check the path and try again.", color='red'))
        ask_for_path()
    except Exception as e:
        print(color_text(f"\nAn error occurred: {e}", color='red'))
        print(color_text("\nThe program will close automatically in 10 seconds...\n"))
        time.sleep(10)
        sys.exit()


if __name__ == "__main__": 

    spreadsheet, file_path = ask_for_path()
    ids_path = file_path.replace(".xlsx", ".ids")

    excel2ids(spreadsheet, ids_path)

    time.sleep(1)
    print(color_text("\nThe program will close automatically in 10 seconds...\n"))
    time.sleep(10)
    sys.exit()