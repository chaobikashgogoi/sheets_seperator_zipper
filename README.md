The program will:

Group the data by the specified column (default Column B).
Create temporary Excel files for each group.
Add these files to a zip archive ("Separated_Data_Archive.zip").
Clean up all temporary files and the "SEPERATED_DATA" folder, ensuring only the zip file remains.


##########
Example:
If your input file C:/Users/YourName/Documents/tea_data.xlsx has data grouped by Column B with values "Estate A", "Estate B", and "" (blank), the program will:

Create temporary Excel files for each group (e.g., "Estate A", "Estate B", "Blank").
Create a zip file C:/Users/YourName/Documents/Separated_Data_Archive.zip containing:
SEPERATED_DATA/Estate A.xlsx
SEPERATED_DATA/Estate B.xlsx
SEPERATED_DATA/Blank.xlsx
Delete all temporary files and the "SEPERATED_DATA" folder.
The only output will be Separated_Data_Archive.zip in the input file's directory.

############
POINTS TO REMEMBER:
EXCEL FILE NAME SHOULD BE: tea_data.xlsx 
