#-------------------------------------------------------------------------------
# Name:        NPMRDS Tool (Step 1)
# Purpose:     Processes NPMRDS data from RITIS.
#
# Author:      Christian Matthews, Rockingham Planning Commission
#              cmatthews@rpc-nh.org
#
# Created:     01/01/2018
# Updated:     01/03/2018
#-------------------------------------------------------------------------------

# Import System libraries
import os
import sys
import glob
import random
import re
import win32com
import arcpy, fnmatch
from win32com.client import Dispatch

scripts_dir = "O:\d-multiyear\d-CongestionManagement\d-tool"
strcode = \
'''
Sub ProcessExcel()
    Dim i As Integer
    Dim xTCount As Variant
    Dim xWs As Worksheet
    Dim xPath As String
    On Error Resume Next
Application.ScreenUpdating = False
LInput:
    xTCount = 3
    Range("B1").Select
    ActiveCell.FormulaR1C1 ="=LEFT(RC[-1],FIND(""index"",RC[-1])+LEN(""index"")-1)"
    xPath = Application.ActiveWorkbook.Path
    Set xWs = ActiveWorkbook.Worksheets.Add(Sheets(1))
    SheetName = Worksheets(2).Range("B1") & Worksheets(2).Range("A2")
    xWs.Name = "Combined"
    Worksheets(2).Range("A3").EntireRow.Copy Destination:=xWs.Range("A1")
    For i = 2 To Worksheets.Count
        Worksheets(i).Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy _
               Destination:=xWs.Cells(xWs.UsedRange.Cells(xWs.UsedRange.Count).Row + 1, 1)
    Next
Worksheets(1).Copy
For Each r In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
    If r = "" Then
        r.Value = 0
        r.NumberFormat = "0.00"
    ElseIf IsNumeric(r) Then
        r.Value = CSng(r.Value)
        r.NumberFormat = "0.00"
    End If
Next
Application.ScreenUpdating = True
Worksheets(1).SaveAs Filename:=xPath & "\\"& SheetName & ".xlsx"
End Sub
'''

x1 = Dispatch("Excel.Application")
x1.Visible = False
x1.DisplayAlerts = False

for script_file in glob.glob(os.path.join(scripts_dir, "*.xml")):
    (file_path, file_name) = os.path.split(script_file)
    objworkbook = x1.Workbooks.Open(script_file)
    print("Processing {}".format(file_name))
    xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(strcode.strip())
    x1.Application.Run("ProcessExcel")
    x1.Workbooks.Close()
    x1.Application.Quit()
    print("Processed file successfully!")
x1.Quit()

arcpy.env.overwriteOutput = True
arcpy.env.workspace = r"O:\d-multiyear\d-CongestionManagement\d-NPMRDSTool"
os.chdir(r"O:\d-multiyear\d-CongestionManagement\d-NPMRDSTool")
print("Imported System Modules")

#Import Tables
tables = arcpy.ListFiles("*.xlsx")
dbPath = r"O:\d-multiyear\d-CongestionManagement\d-NPMRDSTool"
dbName = "NPMRDS"
arcpy.CreateFileGDB_management(dbPath,dbName)
for table in tables:
    if " " in table:
        os.rename(table, table.replace(" ", ""))
tables = arcpy.ListFiles("*.xlsx")
for table in tables:
    if "(" in table:
        os.rename(table, table.replace("(", ""))
tables = arcpy.ListFiles("*.xlsx")
for table in tables:
    if ")" in table:
        os.rename(table, table.replace(")", ""))
tables = arcpy.ListFiles("*.xlsx")
for table in tables:
    arcpy.ExcelToTable_conversion(table, "NPMRDS.gdb\\" + os.path.splitext(os.path.basename(table))[0])
print("Imported Tables")

#Calculate Fields
arcpy.env.workspace = r"O:\d-multiyear\d-CongestionManagement\d-NPMRDSTool\NPMRDS.gdb"
tables = arcpy.ListTables()
for table in tables:
    if fnmatch.fnmatch(table,'*travel*'):
        if 'weekday' in table:
            arcpy.AddField_management(table,"TTI_Peak_AM","DOUBLE")
            tables = arcpy.ListTables()
            arcpy.CalculateField_management(table,"TTI_Peak_AM", "(!F6_00_AM!+!F7_00_AM!+!F8_00_AM!)/3", "PYTHON_9.3")
            tables = arcpy.ListTables()
            arcpy.AddField_management(table,"TTI_Peak_PM","DOUBLE")
            tables = arcpy.ListTables()
            arcpy.CalculateField_management(table,"TTI_Peak_PM", "(!F4_00_PM!+!F5_00_PM!+!F6_00_PM!)/3", "PYTHON_9.3")
            tables = arcpy.ListTables()
            arcpy.AddField_management(table,"TTI_Peak_AVG","DOUBLE")
            tables = arcpy.ListTables()
            arcpy.CalculateField_management(table,"TTI_Peak_AVG", "(!TTI_Peak_AM!+!TTI_Peak_PM!)/2", "PYTHON_9.3")
            tables = arcpy.ListTables()
        else:
            arcpy.AddField_management(table,"TTI_Week_AVG","DOUBLE")
            tables = arcpy.ListTables()
            expression = ("(!F12_00_AM!+!F1_00_AM!+!F2_00_AM!+!F3_00_AM!+!F4_00_AM!+!F5_00_AM!+!F6_00_AM!+!F7_00_AM!+!F8_00_AM!+!F9_00_AM!+!F10_00_AM!+!F11_00_AM!+!F12_00_PM!+!F1_00_PM!+!F2_00_PM!+!F3_00_PM!+!F4_00_PM!+!F5_00_PM!+!F6_00_PM!+!F7_00_PM!+!F8_00_PM!+!F9_00_PM!+!F10_00_PM!+!F11_00_PM!)/24")
            arcpy.CalculateField_management(table,"TTI_Week_AVG", expression)
print("Calculated Fields")

