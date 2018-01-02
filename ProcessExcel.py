# Import System libraries
import os
import sys
import glob
import random
import re
import win32com
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
LInput:
    xTCount = 3
    Range("B1").Select
    ActiveCell.FormulaR1C1 ="=LEFT(RC[-1],FIND(""index"",RC[-1])+LEN(""index"")-1)"
    xPath = Application.ActiveWorkbook.Path
    Set xWs = ActiveWorkbook.Worksheets.Add(Sheets(1))
    xWs.Name = Worksheets(2).Range("B1") & " " & Worksheets(2).Range("A2")
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
Range("AB1").Select
ActiveCell.Formula = "AM_PEAK_6_9"
Range("AC1").Select
ActiveCell.Formula = "PM_PEAK_4_7"
Range("AD1").Select
ActiveCell.Formula = "AVG_PEAK"
Dim x As Integer
      NumRows = Range("AA1", Range("AA1").End(xlDown)).Rows.Count
      Range("AB2").Select
      xRow = 2
      xColumn = 2
      For x = 2 To NumRows
         ActiveCell.Formula = "=AVERAGE(J" & xRow & ":L" & xColumn & ")"
         ActiveCell.Offset(1, 0).Select
         xRow = xRow + 1
         xColumn = xColumn + 1
      Next
      Range("AC2").Select
      xRow = 2
      xColumn = 2
      For x = 2 To NumRows
         ActiveCell.Formula = "=AVERAGE(T" & xRow & ":V" & xColumn & ")"
         ActiveCell.Offset(1, 0).Select
         xRow = xRow + 1
         xColumn = xColumn + 1
      Next
      Range("AD2").Select
      xRow = 2
      xColumn = 2
      For x = 2 To NumRows
         ActiveCell.Formula = "=AVERAGE(AB" & xRow & ":AC" & xColumn & ")"
         ActiveCell.Offset(1, 0).Select
         xRow = xRow + 1
         xColumn = xColumn + 1
      Next
Worksheets(1).SaveAs Filename:=xPath & "\\"& xWs.Name & ".xlsx"
End Sub
'''

x1 = Dispatch("Excel.Application")
x1.Visible = False 
x1.DisplayAlerts = False 

for script_file in glob.glob(os.path.join(scripts_dir, "*.xml")):
    (file_path, file_name) = os.path.split(script_file)
    objworkbook = x1.Workbooks.Open(script_file)
    xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(strcode.strip())
    x1.Application.Run("ProcessExcel")
    x1.Workbooks.Close()
    x1.Application.Quit()
    print("Macro ran successfully!")
    
x1.Quit()
