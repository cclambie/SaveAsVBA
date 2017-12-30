################
Project: Save As VBA macros for Excel
Purpose: To convert CSV or simple Data in Excel sheet to either CSV or OFX, put simple, easily importable data for WaveApps or other accounting software

Version: 0.1
Date written: 30 Dec 2017

Author: Craig Lambie
Author URI: craiglambie.com

Data:
Generally expects data to be in 3 columns: Date, Description and Amount - could be fiddled with to accommodate more, that is what I have tested

Routines: 
SaveAsCSV - saves the current active sheet as a CSV file with the same name as the open file, keeping the data formats
A fork of the code found on this StackOverflow.com question https://stackoverflow.com/questions/37037934/excel-macro-to-export-worksheet-as-csv-file-without-leaving-my-current-excel-sh/37038840

SaveAsOFX - saves the current active sheet as a OFX file with the same name as the open file, with option to save as Credit Card data (the purpose I designed it for) or not
As of version 0.1, this is pretty new.
Borrows heavily on the XLS2OFX Converter v1.0 by Josep Bori, but with many updates and changes, and needs a bunch more, if I have time.
Inspired by the free online converter that takes to long: http://csvconverter.biz/

Usage:
1. Open your Excel file (of any type)
2. Push the data you want to a clean sheet with 3 headers/ top row
3. Go to ribbon View>Macros
4. On dialogue box, dropdown box "Macros in:" select "All Workbooks"
5. Select SaveAsCSV or SaveAsOFX as required
6. Visit the location your excel file is saved to find the CSV or OFX file
7. Done :) 

Installation:
1. Open Excel
2. Open VBA (Alt + F11)
3. Import the .cls file
