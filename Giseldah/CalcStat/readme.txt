Stat definition file statdata.csv:

You can view this comma separated values file with something like a spreadsheet application. Upload it to Google Sheets for example or use LibreOffice Calc, Microsoft Excel etc.

Compilation of statdata.csv:

Execute statdcmp.vbs (double click) to generate the CalcStat script file. This will contain the major calculation function.
Output version "Percentage calculations only" supports ratings<->percentages and rating penetration calculations and will do for all known Lotro plugins at this time.

Installation/use:

Lotro Plugins:
Extract the zip file to your Lotro Plugins directory. After, you should have a directory like "C:\Users\UserName\Documents\The Lord of the Rings Online\Plugins\Giseldah\CalcStat"
Plugins can import this library using the statement: import Giseldah.CalcStat
The directory already contains a "Percentage calculations only" version by default, so you don't need to compile statdata.csv if this is all you require.

OpenOffice/LibreOffice:
Compile statdata.csv with StarBasic as output script type. This will generate the CalcStat.bas file with the CalcStat function.
Inside the Calc spreadsheet application, create a new CalcStat and CalcSsup macro module, either under "My macros & Dialogs" for use globally in any spreadsheet or create the modules in a specific spreadsheet.
Copy the content of the .bas files to these new modules.

VB-script (Windows):
Compile statdata.csv with VB-script as output script type. This will generate the CalcStat.vbs file with the CalcStat function.
You can include CalcStat.vbs and CalcSsup.vbs conveniently in a script job by using a WSF file (see https://en.wikipedia.org/wiki/Windows_Script_File).
See the vbs_example directory for an example script.
