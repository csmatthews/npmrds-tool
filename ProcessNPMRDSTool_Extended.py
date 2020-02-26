# -----------------------------------------------------------------------------
# Name:        NPMRDS Tool
# Purpose:     Processes NPMRDS data from RITIS.
#
# Author:      Christian Matthews, Rockingham Planning Commission
#              cmatthews@therpc.org
# -----------------------------------------------------------------------------

# Import System libraries
import os
import glob
import win32com
import arcpy
import fnmatch
import numpy
from win32com.client import Dispatch
print("Imported libraries")

# Set User Paths
directory = r"O:\d-multiyear\d-CongestionManagement\d-data\d-2019"
gdbName = "NPMRDS.gdb"


# Create code for VBA
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
    If IsNumeric(r) Then
        r.NumberFormat = "0.00"
        r.Value = CDec(r.Value)
    End If
Next
Application.ScreenUpdating = True
Worksheets(1).SaveAs Filename:=xPath & "\\"& SheetName & ".xlsx"
End Sub
'''

# Setup Excel Parameters
x1 = Dispatch("Excel.Application")
x1.Visible = False
x1.DisplayAlerts = False
print("Set Excel parameters")

# Apply VBA code
for script_file in glob.glob(os.path.join(directory, "*.xml")):
    (file_path, file_name) = os.path.split(script_file)
    objworkbook = x1.Workbooks.Open(script_file)
    print("Processing {}".format(file_name))
    xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(strcode.strip())
    x1.Application.Run("ProcessExcel")
    x1.Workbooks.Close()
    x1.Application.Quit()
    print("Processed {} successfully".format(file_name))
x1.Quit()
print("Finished processing .XML files")

# Set ArcGIS Parameters
arcpy.env.overwriteOutput = True
arcpy.env.workspace = directory
os.chdir(directory)
print("Set ArcGIS Parameters")

# Import Tables
excelFiles = arcpy.ListFiles("*.xlsx")
arcpy.CreateFileGDB_management(directory, gdbName)
for excel in excelFiles:
    os.rename(excel, excel.replace(" ", "").replace("(", "").replace(")", ""))
excelFiles = arcpy.ListFiles("*.xlsx")
for excel in excelFiles:
    arcpy.ExcelToTable_conversion(excel,
                                  gdbName+"\\" +
                                  os.path.splitext(os.path.basename(excel))[0])
print("Imported Tables")

# Calculate Fields
#arcpy.env.workspace = directory + "\\" + gdbName
#tables = arcpy.ListTables()
#peakAM = ["F6_00_AM", "F6_15_AM", "F6_30_AM", "F6_45_AM", "F7_00_AM",
#          "F7_15_AM", "F7_30_AM", "F7_45_AM", "F8_00_AM", "F8_15_AM",
#          "F8_30_AM", "F8_45_AM"]
#for table in tables:
#    if fnmatch.fnmatch(table, '*travel*'):
#        if 'weekday' in table:
#            arcpy.AddField_management(table, "TTI_Peak_AM", "DOUBLE")
#            arcpy.AddField_management(table, "TTI_Peak_PM", "DOUBLE")
#            with arcpy.da.Update(table, peakAM):
#            arcpy.CalculateField_management(table, "TTI_Peak_AM",
#                                            "numpy.mean(!F6_00_AM!,!F6_15_AM!,!F6_30_AM!,!F6_45_AM!,!F7_00_AM!,!F7_15_AM!,!F7_30_AM!,!F7_45_AM!,!F8_00_AM!,!F8_15_AM!,!F8_30_AM!,!F8_45_AM!)",
#                                            "PYTHON_9.3")
#            arcpy.CalculateField_management(table, "TTI_Peak_PM",
#                                            "numpy.mean(!F4_00_PM!,!F4_15_PM!,!F4_30_PM!,!F4_45_PM!,!F5_00_PM!,!F5_15_PM!,!F5_30_PM!,!F5_45_PM!,!F6_00_PM!,!F6_15_PM!,!F6_30_PM!,!F6_45_PM!)",
#                                            "PYTHON_9.3")
#    if fnmatch.fnmatch(table, '*buffer*'):
#        if 'weekday' in table:
#            arcpy.AddField_management(table, "BI_Peak_AM", "DOUBLE")
##            tables = arcpy.ListTables()
#            arcpy.CalculateField_management(table, "BI_Peak_AM",
#                                            "numpy.mean(!F6_00_AM!,!F6_15_AM!,!F6_30_AM!,!F6_45_AM!,!F7_00_AM!,!F7_15_AM!,!F7_30_AM!,!F7_45_AM!,!F8_00_AM!,!F8_15_AM!,!F8_30_AM!,!F8_45_AM!)",
#                                            "PYTHON_9.3")
##            tables = arcpy.ListTables()
#            arcpy.AddField_management(table, "BI_Peak_PM", "DOUBLE")
##            tables = arcpy.ListTables()
#            arcpy.CalculateField_management(table, "BI_Peak_PM",
#                                            "numpy.mean(!F4_00_PM!,!F4_15_PM!,!F4_30_PM!,!F4_45_PM!,!F5_00_PM!,!F5_15_PM!,!F5_30_PM!,!F5_45_PM!,!F6_00_PM!,!F6_15_PM!,!F6_30_PM!,!F6_45_PM!)",
#                                            "PYTHON_9.3")
##            tables = arcpy.ListTables()
#    if fnmatch.fnmatch(table, '*planning*'):
#        if 'weekday' in table:
#            arcpy.AddField_management(table, "PTI_Peak_AM", "DOUBLE")
##            tables = arcpy.ListTables()
#            arcpy.CalculateField_management(table, "PTI_Peak_AM",
#                                            "numpy.mean(!F6_00_AM!,!F6_15_AM!,!F6_30_AM!,!F6_45_AM!,!F7_00_AM!,!F7_15_AM!,!F7_30_AM!,!F7_45_AM!,!F8_00_AM!,!F8_15_AM!,!F8_30_AM!,!F8_45_AM!)",
#                                            "PYTHON_9.3")
##            tables = arcpy.ListTables()
#            arcpy.AddField_management(table, "PTI_Peak_PM", "DOUBLE")
##            tables = arcpy.ListTables()
#            arcpy.CalculateField_management(table, "PTI_Peak_PM",
#                                            "numpy.mean(!F4_00_PM!,!F4_15_PM!,!F4_30_PM!,!F4_45_PM!,!F5_00_PM!,!F5_15_PM!,!F5_30_PM!,!F5_45_PM!,!F6_00_PM!,!F6_15_PM!,!F6_30_PM!,!F6_45_PM!)",
#                                            "PYTHON_9.3")
##            tables = arcpy.ListTables()
#            arcpy.AddField_management(table, "PTI_Peak_AVG", "DOUBLE")
# #           tables = arcpy.ListTables()
print("Calculated Fields")