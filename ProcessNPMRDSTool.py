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
from win32com.client import Dispatch
print("Imported libraries")

# Set User Paths
directory = r""

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
