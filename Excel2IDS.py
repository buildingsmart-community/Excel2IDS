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

# TEST:
# TODO regexes everywhere with "..." --> TEST
# TODO discipline in column --> TEST
# TODO split enums with new line or comma with or without spaces

# Backlog:
# TODO cardinality "APL_CARDINAL": 1,
# TODO BONUS: replace "REPLACEME" replaces with "X"
# TODO new template
# TODO no settings - recognize the row/column
# TODO warn on common errors like no predefined types in enums
# TODO expose, currently hardcoded:	    "IFC_VERSION_DEFAULT": "IFC4X3_ADD2",
# TODO expose, currently hardcoded:		"IDS_TITLE": "H2",
# TODO expose, currently hardcoded:		"IDS_AUTHOR": "H3",
# TODO expose, currently hardcoded:		"IDS_DATE": "H4",
# TODO expose, currently hardcoded:		"IDS_VERSION": "H5",
# TODO expose, currently hardcoded:		"IDS_COPYRIGHT": "H6",
# TODO expose, currently hardcoded:		"IFC_VERSION": "H7",
# TODO expose, currently hardcoded:		"IFC_VERSION_DEFAULT": "IFC4X3_ADD2",
# TODO expose, currently hardcoded:		"IDS_DESCRIPTION": "H8",
# TODO expose, currently hardcoded:		"SPE_DESCR": 1,
# TODO expose, currently hardcoded:		"SPE_INSTR": 1,
# TODO expose, currently hardcoded:		"SPE_IDENT": 1,


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
    start_row = sheet[s.DEFAULT_START_CELL].row
    start_col = sheet[s.DEFAULT_START_CELL].column
    end_row = sheet.max_row + 1
    end_col = sheet.max_column + 1

    for col in tqdm(range(start_col, end_col), desc="Processing Excel columns."):
        column_letter = openpyxl.utils.get_column_letter(col)
        specification_name = sheet[f"{column_letter}{s.SPE_NAME}"].value

        if sheet[f'{column_letter}{s.APL_INCLUDE}'].value:
            ### add applicability
            applicability = []
            # add entity and predefined type
            entity_cell = process_value(sheet[f'{column_letter}{s.APL_ENTITY}'].value)
            predefined_type_cell = process_value(sheet[f'{column_letter}{s.APL_PRED_TYPE}'].value)
            if entity_cell:
                if predefined_type_cell:
                    entity = ids.Entity(name=entity_cell, predefinedType=predefined_type_cell)
                else:
                    entity = ids.Entity(name=entity_cell)
                # if '.' in entity_cell:
                #     entity = ids.Entity(name=entity_cell.split('.')[0].upper(), predefinedType=entity_cell.split('.')[1].upper())
                # else:
                applicability.append(entity)
            # add property
            if sheet[f'{column_letter}{s.APL_PNAME}'].value:
                property = ids.Property(
                    propertySet=process_value(sheet[f'{column_letter}{s.APL_PSET}'].value), 
                    baseName=process_value(sheet[f'{column_letter}{s.APL_PNAME}'].value),
                    dataType=process_value(sheet[f'{column_letter}{s.APL_PDTYPE}'].value),
                    cardinality="required", # TODO should be optional(?)
                    )
                pv = sheet[f"{column_letter}{s.APL_PVAL}"].value
                if not_empty(pv):
                    property.value = process_value(pv)
                    property.instructions = f"All objects with code like: '{pv}'"
                applicability.append(property)
            # add classification
            if not_empty(sheet[f'{column_letter}{s.APL_CLASS_SYS}'].value):
                classification = ids.Classification(
                    system=process_value(sheet[f'{column_letter}{s.APL_CLASS_SYS}'].value),
                    value=process_value(sheet[f'{column_letter}{s.APL_CLASS_CODE}'].value))
                applicability.append(classification)
            # add attribute
            if not_empty(sheet[f'{column_letter}{s.APL_ANAME}'].value):
                attribute = ids.Attribute(
                    name=process_value(sheet[f'{column_letter}{s.APL_ANAME}'].value),
                    value=process_value(sheet[f'{column_letter}{s.APL_AVALUE}'].value)
                )
                applicability.append(attribute)
            # add material
            if not_empty(sheet[f'{column_letter}{s.APL_MATERIAL}'].value):
                material = ids.Material(
                    value=process_value(sheet[f'{column_letter}{s.APL_MATERIAL}'].value)
                )
                applicability.append(material)

            ### add requirement(s)
            requirements = []
            for row in range(start_row, end_row):
                if sheet[f'{s.REQ_INCLUDE}{row}'].value:

                    cell_value = sheet.cell(row=row, column=col).value
                    if cell_value in ["X", "x"]:
                        # add entity and predefined type
                        entity_cell = process_value(sheet[f'{s.REQ_ENTITY}{row}'].value)
                        predefined_type_cell = process_value(sheet[f'{s.REQ_PRED_TYPE}{row}'].value)
                        if entity_cell:
                            if predefined_type_cell:
                                entity = ids.Entity(name=entity_cell, predefinedType=predefined_type_cell)
                            else:
                                entity = ids.Entity(name=entity_cell)
                            requirements.append(entity)
                        # add property
                        if not_empty(sheet[f'{s.REQ_PNAME}{row}'].value):
                            property = ids.Property(
                                propertySet=process_value(sheet[f'{s.REQ_PSET}{row}'].value), 
                                baseName=process_value(sheet[f'{s.REQ_PNAME}{row}'].value),
                                dataType=process_value(sheet[f'{s.REQ_PDTYPE}{row}'].value),
                                cardinality="required"
                            )
                            if not_empty(sheet[f'{s.REQ_PVAL}{row}'].value):
                                property.value = process_value(sheet[f'{s.REQ_PVAL}{row}'].value)
                            if not_empty(sheet[f'{s.REQ_URI}{row}'].value):
                                property.uri = process_value(sheet[f'{s.REQ_URI}{row}'].value)
                            requirements.append(property)
                        # add classification
                        if not_empty(sheet[f'{s.REQ_CLASS_SYS}{row}'].value):
                            classification = ids.Classification(
                                system=process_value(sheet[f'{s.REQ_CLASS_SYS}{row}'].value),
                                value=process_value(sheet[f'{s.REQ_CLASS_CODE}{row}'].value)
                            )
                            requirements.append(classification)
                        # add attribute
                        if not_empty(sheet[f'{s.REQ_ANAME}{row}'].value):
                            attribute = ids.Attribute(
                                name=process_value(sheet[f'{s.REQ_ANAME}{row}'].value),
                                value=process_value(sheet[f'{s.REQ_AVALUE}{row}'].value)
                            )
                            applicability.append(attribute)
                        # add material
                        if not_empty(sheet[f'{s.REQ_MATERIAL}{row}'].value):
                            material = ids.Material(
                                value=process_value(sheet[f'{s.REQ_MATERIAL}{row}'].value)
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
                        # TODO process 'REPLACEME'
                        # TEMP workaround, assuming if not X it must be entity!:                    
                        if '.' in cell_value:
                            entity = ids.Entity(name=cell_value.split('.')[0], predefinedType=cell_value.split('.')[1])
                        else:
                            entity = ids.Entity(name=process_value(cell_value))
                        requirements.append(entity)
                    else:
                        # skip an empty cell
                        pass

            if requirements:
                disciplines = split_multivalue(sheet[f"{column_letter}{s.APL_DISCIPLINE}"].value)
                ifc_version = sheet[s.IFC_VERSION].value
                if not ifc_version in ['IFC2X3','IFC4','IFC4X3_ADD2']:
                    ifc_version = s.IFC_VERSION_DEFAULT
                for discipline in disciplines:
                    add_to_ids(
                        discipline,
                        specification_name,
                        applicability,
                        requirements,
                        purpose=discipline,
                        #TODO process phases, below TEMP workarond with fixed milestone
                        milestone=MILESTONE,
                        ifc_version=ifc_version
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
    ifc_version="IFC4X3_ADD2",
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
            milestone=milestone
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
            ifcVersion=ifc_version,
            # identifier=str(sheet[f'{REQ_ID}{row}'].value.strip()),
        )
        new_spec.applicability = copy.deepcopy(applicability)
        new_spec.requirements = copy.deepcopy(requirements)
        ids_list[ids_name].specifications.append(new_spec)


QUOTED_PATTERN = r'^".*"$'

def process_value(cell_value):
    """ Look at the value and convert to enumeration of pattern if needed """
    if not_empty(cell_value):
        # remove trailing spaces (unless non-string)
        if isinstance(cell_value, str):
            cell_value=cell_value.strip()
            # if there are multiple lines in a single cell, split it into enumeration of literal values
            if "\n" in cell_value or "," in cell_value or ";" in cell_value :
                cell_value = ids.Restriction(options={"enumeration": split_multivalue(cell_value)})
            # if it's in quotation, turn into pattern restriction (regex):
            elif bool(re.match(QUOTED_PATTERN, cell_value)):
                cell_value = ids.Restriction(options={"pattern": cell_value[1:-1]})
            # in other cases, return as is (empty or literal value)
        elif isinstance(cell_value, bool):
            cell_value=str(cell_value)
    return cell_value


SPLIT_PATTERN=r'\s*[,;\n]\s*'

def split_multivalue(text):
    """ Split discipline/phase text if plural (allowed multiline, comma-separated and semicolon-separated).  """
    return re.split(SPLIT_PATTERN, text)


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

    MILESTONE = 'LOD400' #TODO TEMP workaround for phases

    excel2ids(spreadsheet, ids_path)

    time.sleep(1)
    print(color_text("\nThe program will close automatically in 10 seconds...\n"))
    time.sleep(10)
    sys.exit()