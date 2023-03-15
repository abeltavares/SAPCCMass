[![Made with VBA](https://img.shields.io/badge/Made%20with-VBA-blue)](https://docs.microsoft.com/en-us/dotnet/visual-basic/)
[![SAP Automation](https://img.shields.io/badge/SAP-Automation-orange)](https://www.sap.com/products/intelligent-automation.html)

# SAPCCMass

![SAPCCMass](SAPCCMass.png)

SAPCCMass is a tool for automating mass maintenance of cost centers in SAP using transaction codes KS01 and KS02. This tool allows you to create or modify multiple cost centers at once by reading the data from an Excel worksheet and sending the updates to SAP.

## Requirements
- Microsoft Excel 2010 or later.
- SAP GUI for Windows installed on your computer.
- The SAP system that you are working with must be accessible from your computer.

## Files
This repository contains the following files:
- `KS01.bas`: Visual Basic script for mass maintenance of cost centers using tcode KS01 in SAP.
- `KS02.bas`: Visual Basic script for mass maintenance of cost centers using tcode KS02 in SAP.
- `config.bas`: Visual Basic script for configuring the Excel interface settings.
- `run.bas`: Visual Basic script for calling the appropriate subroutine based on the transaction code selected.
- `SAPCCMass.xlsm`: Excel file containing the functional script for mass maintenance of cost centers in SAP.

## Mandatory Fields
### KS01
- Valid From Date
- Valid To Date
- Name
- Person Responsible
- Cost Center Category
- Company Code
- Profit Center
- Functional Area

### KS02
- Valid From Date
- Valid To Date
- Name
- Person Responsible
- Cost Center Category
- Company Code
- Profit Center
- Functional Area

## Usage
To use this script, follow these steps:
1. Open the Excel file "SAPCCMass.xlsm".
2. Fill out the System Name and Cost Center fields in the first row of the worksheet.
3. Choose the tcode (KS01 or KS02).
4. Fill out the fields for each cost center in the respective columns.
5. Click the "Run Script" button.
6. Wait for the script to finish running. A message box will appear when the script has finished processing all cost centers.
7. Check the log in column A for any errors or warnings.

## Copyright
© 2023 Abel Tavares
