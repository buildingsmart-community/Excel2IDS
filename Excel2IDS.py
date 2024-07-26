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


class Settings:
    def __init__(self, settings_file):
        with open(settings_file, 'r') as file:
            settings_dict = json.load(file)
        for key, value in settings_dict.items():
            setattr(self, key, value)


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
        specification_name = sheet[f"{column_letter}{s.SPE_NAME}"].value

        if sheet[f'{column_letter}{s.APL_INCLUDE}'].value == 1:
            ### add applicability
            applicability = []
            # TODO add entity and predefined type
            entity_cell = sheet[f'{column_letter}{s.APL_ENTITY}'].value
            if entity_cell:
                if '.' in entity_cell:
                    entity = ids.Entity(name=entity_cell.split('.')[0].upper(), predefinedType=entity_cell.split('.')[1].upper())
                else: 
                    entity = ids.Entity(name=process_value(entity_cell.upper()))
                applicability.append(entity)
            # add property
            if sheet[f'{column_letter}{s.APL_PNAME}'].value:
                property = ids.Property(
                    propertySet=sheet[f'{column_letter}{s.APL_PSET}'].value.strip(), 
                    baseName=sheet[f'{column_letter}{s.APL_PNAME}'].value.strip(),
                    cardinality="required", # TODO should be optional(?)
                    dataType='IFCLABEL',
                    )
                pv = sheet[f"{column_letter}{s.APL_PVAL}"].value
                if not_empty(pv):
                    property.value = process_value(pv.strip())
                    property.instructions = f"All objects with code like: '{pv.strip()}'"
                applicability.append(property)
            # add classification
            if sheet[f'{column_letter}{s.APL_CLASS_SYS}'].value:
                classification = ids.Classification(
                    system=sheet[f'{column_letter}{s.APL_CLASS_SYS}'].value.strip(),
                    value=process_value(sheet[f'{column_letter}{s.APL_CLASS_CODE}'].value.strip())
                )
                applicability.append(classification)
            # add material
            if sheet[f'{column_letter}{s.APL_MATERIAL}'].value:
                material = ids.Material(
                    value=process_value(sheet[f'{column_letter}{s.APL_MATERIAL}'].value.strip())
                )
                applicability.append(material)

            for row in range(start_row, end_row):
                if sheet[f'{s.REQ_INCLUDE}{row}'].value == 1:

                    cell_value = sheet.cell(row=row, column=col).value
                    ### add requirement(s)
                    requirements = []
                    if cell_value in ["X", "x"]:
                        # add property
                        if not_empty(sheet[f'{s.REQ_PNAME}{row}'].value):
                            property = ids.Property(
                            propertySet=sheet[f'{s.REQ_PSET}{row}'].value.strip(), 
                            baseName=sheet[f'{s.REQ_PNAME}{row}'].value.strip(),
                            dataType=sheet[f'{s.REQ_DTYPE}{row}'].value.strip().upper(),
                            cardinality="required"
                            )
                            if not_empty(sheet[f'{s.REQ_PVAL}{row}'].value):
                                property.value = sheet[f'{s.REQ_PVAL}{row}'].value.strip()
                            elif not_empty(sheet[f'{s.REQ_PVAL_PATTERN}{row}'].value):
                                property.value = process_value(sheet[f'{s.REQ_PVAL_PATTERN}{row}'].value.strip())
                                # TODO TEMP if isinstance(property.value, str):
                                # TODO TEMP property.value = ids.Restriction(options={"pattern": sheet[f'{REQ_PVAL_PATTERN}{row}'].value})
                            if sheet[f'{s.REQ_URI}{row}'].value:
                                property.uri = sheet[f'{s.REQ_URI}{row}'].value.strip()
                            requirements.append(property)
                        # add classification
                        if sheet[f'{s.REQ_CLASS_SYS}{row}'].value:
                            classification = ids.Classification(
                                system=sheet[f'{s.REQ_CLASS_SYS}{row}'].value.strip(),
                                value=process_value(sheet[f'{s.REQ_CLASS_CODE}{row}'].value.strip())
                            )
                            requirements.append(classification)
                        # add material
                        if sheet[f'{s.REQ_MATERIAL}{row}'].value:
                            material = ids.Material(
                                value=process_value(sheet[f'{s.REQ_MATERIAL}{row}'].value.strip())
                            )
                            requirements.append(material)
                        # TODO add entity
                        # entity_cell = sheet[f'{column_letter}{s.REQ_ENTITY}'].value
                        # if entity_cell:
                        #     if '.' in entity_cell:
                        #         entity = ids.Entity(name=entity_cell.split('.')[0].upper(), predefinedType=entity_cell.split('.')[1].upper())
                        #     else:
                        #         entity = ids.Entity(name=process_value(entity_cell.upper()))
                        #     requirements.append(entity)
                    # ids.Entity(name=sheet[f'{column_letter}{REQ_IFC}'].value.upper())
                    elif not_empty(cell_value):
                        # assuming if a value is not 'x' it is list of possible entities
                        if '.' in cell_value:
                            entity = ids.Entity(name=cell_value.split('.')[0].upper(), predefinedType=cell_value.split('.')[1].upper())
                        else:
                            entity = ids.Entity(name=process_value(cell_value.upper()))
                        requirements.append(entity)
                    else:
                        # skip an empty cell
                        pass

                    if requirements:
                        disciplines = sheet[f"{col}{s.APP}"].value.split(",")
                        for discipline in disciplines:
                            add_to_ids(
                                discipline,
                                specification_name,
                                applicability,
                                requirements,
                                purpose=discipline,
                                milestone=s.MILESTONE,
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
            ifcVersion=s.IFC_VERSION,
            # identifier=str(sheet[f'{REQ_ID}{row}'].value.strip()),
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

    s = Settings('settings.json')
    ids_list = {}

    spreadsheet, file_path = ask_for_path()
    ids_path = file_path.replace(".xlsx", ".ids")

    excel2ids(spreadsheet, ids_path)

    time.sleep(1)
    print(color_text("\nThe program will close automatically in 10 seconds...\n"))
    time.sleep(10)
    sys.exit()