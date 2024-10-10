# piping_loads
Read load data from a spreadsheet and write to a file in SACS format.

## Load Case ID
Input list of load cases:

 - Sheet: Worksheet name
 - Column: Index of first column of data (e.g. for column D enter 4)
 - LOADCN: Load case ID
 - LOADLB: Load case description
 - LOAD_ID: Remark at the end of the LOAD line
 - SUFFIX: If LOAD_ID is PSXX then the suffix is added to the remark, e.g. LOAD_ID=PS03, SUFFIX=90, remark=PS03_90. If total width is greater than 8 characters it will be trucated to 8.

## Data sheets
Joint labels are assumed to be in the first column, rows 3 to 21.
Data starts in column specified on Load Case ID sheet, rows 3 to 21, 6 columns wide.

![alt](img/fig1.png)
