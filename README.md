# PayoutAutomator

## Prequisites

The only prequisite external installation you need for this program will be to have a Python instance on your computer. The documentation is stated here:
https://www.python.org/downloads/

You may need to do some Python installations if you do not have the packages:
`!pip install openpyxl
!pip install pandas`

## How to Run

This program is a 2 step CLI program at the moment and will require you to have multiple files for digestion at runtime.

1) Run preprocess.py

You will need multiple files to input to this program
- YNE_Sales.csv
- YNM_Sales.csv

You will receive multiple files as an output to this step
- YNE_Sales_Final.csv
- YNM_Sales_Final.csv
- No_Cost_Items.csv

The No_Cost_Items.csv contains items that have no costs assigned to them. These should be resolved before going to the second step. Any item that is not assigned to a cost shall not be ready for digestion in the main part of the
program and shall be placed in the input file "Pending_Rollovers.csv" that is needed for the 2nd part of the program.

It is recommended that you consult a member in Finance or Quality Control first before proceeding to the 2nd step.

2) Run main.py

You will need multiple files to input to this program
- YNE_Sales_Final.csv
- YNM_Sales_Final.csv
- Pending_Rollovers.csv
- Rollovers.csv

You will receive a single output file from this program
- YNM Payout Prototype.xlsx
