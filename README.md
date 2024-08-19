## Use
This code is designed to process a Bill of Materials (BOM) Excel file and match it with corresponding STEP files in a 
specified directory. The main function, execute_func, reads the BOM file, checks its format, and searches for STEP 
files in the given directory. For each matching STEP file, it calculates the thickness and updates the BOM with this 
information. The updated BOM is then saved as a new Excel file. Helper functions assist in finding STEP files, reading 
the BOM, checking its format, and calculating thickness from the STEP files. The code also handles errors and updates 
the user interface with the status of the operation.