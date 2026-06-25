# Automated-inspection-report-system
This is an automated inspection report system that converts a .csv file to an .xlsx file along with summary and reports.
This project marks the week 3 of python and it makes use of numpy, pandas, openpyxl and matplotlib to generate a excel based report using a CSV(Comma Separated Values) file. It takes input as the CSV file, so user are expected to add the location and the name of the CSV file and then the save location of the output files. It then creates a .xlsx file called the Automatic Inspection Report and 2 png file which is embeded into the report in the Summary stats sheet. The .xlsx file consists of 4 sheets the first one being the raw data converted from the CSV file to the .xlsx file, second sheet is the summary sheet that provides with the deviation, the status based on deviation and the Flagged status based on the standard deviation. Third sheet is the Failed only sheet consisting of the details of the items that failed and finally forth sheet is the Inspector Summary sheet providing details about the inspector that inspected the respective parts.

**Important**
The csv file is supposed to have these headers in the same order as displayed below:
Part ID, Dimension Measured, Nominal Value, Actual Value, Tolerance, Inspector
