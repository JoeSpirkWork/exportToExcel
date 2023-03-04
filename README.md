This report attempts to give the user an overview of the Denver Drainage Program, written for RS&H by Joseph Spirk and Thomas Hildebrand.
Part 1: Program Overview

The aim of this program is to replace an existing, legacy Drainage Spreadsheet written in VBA code with a more modern visual studio program written in C# using the .NET Framework version 4.6.2. 
Terms:

Dll – Dynamically Linked Library – this is a module that contains functions that can be used by another module (application or .dll). The application we’re using is contained within a .dll file. 
C# - Object Oriented Programming language used to crease cross platform applications. Supported by Microsoft on the .NET Framework
.NET Framework – proprietary software framework developed by Microsoft

Part 2: Loading the Program

The program can be loaded into Microstation Connect or OpenRoads Connect through the following steps:

1)	Find the following folder on your machine view windows file explorer: 
a.	For Microstation Connect: C:\Program Files\Bentley\MicroStation CONNECT Edition\MicroStation\Mdlapps  
b.	For OpenRoads Designer: C:\Program Files\Bentley\OpenRoads Designer CE 10.10\OpenRoadsDesigner\Mdlapps

2)	Ensure you have the latest .dll on your hard drive. Remember where it is. 

3)	Copy the .dll file into the folders from step 1.

4)	Open Microstation or OpenRoads

5)	Bring up the key in command menu: 

 

6)	Type in “mdl load”

7)	Click the “Browse…” button on the lower right

8)	Navigate to the folder from step 1 (if not there already) and double-click “exportToExcel.dll” 

a.	Note: if you do not see .dll files, ensure that on the lower right, “All Files (*.*)” is selected on the file filter.
9)	At this point, the program should be loaded.
Part 2: Running the Excel and element selection portions of the program

1)	In the key in command window, type in “exporttoexcel beginexport”
a.	Note: the key in should begin to auto populate the commands list. If you do not see this, please try to load the program again or contact Joe Spirk (Jabber, or Email – Joe.Spirk@rsandh.com)

2)	A pop up window will come up, on the pop up window coordinate the working folder your excel file will live in and type the name of the excel file (do not include a .xlsx on the end of the file name)
a.	Note: Project Wise folder functionality has not been added to this program yet, we will aim to do this in the future. 

3)	Once all fields are set, click “Create Excel File” on the right hand side. At this point an excel file will come up with headers auto-populated
a.	Note: The headers are from the existing VBA samples I was given. If there is data/headers that your group does not use, feel free to exclude them. 

4)	Next, select all of the items you want to export
a.	Note, right now, the program supports shapes, lines, and linestrings. 

5)	Click “Export”

6)	All of the pertinent data from the select items should be exported to the excel file. 

Last Section: Running Issues
1)	When we select a line string and process it, the middle of string most likely does not fall on the string but instead is out in space. For the VBA program, how were line strings processed? Each line segment at a time? Or 1 overall line string?
2)	We will need to add more robustness within the program. Right now, it is easy to mess up the excel file portion by accidentally clicking create excel file twice. 
3)	It might be nice to give this program the ability to dock within microstation. 
4)	Give the program the ability to work on projectwise


