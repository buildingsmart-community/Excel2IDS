<img src="https://github.com/user-attachments/assets/88e33c7a-e7d7-4d5b-b70f-2a9647c47e7b" alt="drawing" width="150"/>

# Excel2IDS
An Excel template and tool to help generate IDS specifications.

1. Download the latest release of the tool (.exe, .json. and .xlsx files):
<img src="https://github.com/user-attachments/assets/496f8e54-ddce-473d-9278-e2f93e212f2b" alt="drawing" width="400"/>


2. Fill in the Excel file with your specifications (instructions inside)
![image](https://github.com/user-attachments/assets/31782fd4-6bb4-49b5-86a3-7ee93150114e)


3. Run the .exe tool and paste the path to the Excel file. The tool will generate as many IDS files as 'purposes'/'disciplines' found in the file, and save them in the same folder as the Excel file. 
![Excel2IDS_animation](https://github.com/user-attachments/assets/b6bfc2f0-bde7-4951-8a94-471ef6fdb9bc)

# Release notes
Version 0.9.4 supports:
- IDS version 1.0.0
- cardinality
- patterns
- enumerations
- 'REPLACEME' feature
doesn't support:
- reading IDS files
- specifying 'Milestone' (phase) attribute
- using PartOf facet
- combining multiple properties/facets in a single specification

The tool is using the [IfcTester](https://github.com/IfcOpenShell/IfcOpenShell/tree/v0.8.0/src/ifctester) of [IfcOpenShell](https://github.com/IfcOpenShell/IfcOpenShell) ❤️
