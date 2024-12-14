# VBA_MACROS
I have used 3 sample datasets, "Agency Commissions TEST SAMPLE", "Agency Bonus", "Agency Charges". I have written VBA code and record macros to extract data from these 3 datasets to my worksheets, create history, create button for generate PDF file, format blank spaces and generate update data from sources.

2 useful VBA code script given below:


1. VBA code to generate data from the sources:

Sub ImportExcelFiles()
    Dim ws As Worksheet
    Dim excelFilePath As String
    Dim fileNames As Variant
    Dim headings As Variant
    Dim importedWorkbook As Workbook
    Dim currentRow As Long
    Dim lastCol As Integer
    Dim totalCommissions As Double
    Dim totalBonus As Double
    Dim totalChargebacks As Double
    Dim deductions As Double
    Dim totalDeductions As Double
    Dim netPay As Double

    ' Define the file path and file names
    excelFilePath = "C:\Users\BdCalling\Documents\VBA New Project\"
    fileNames = Array("Agency Commissions TEST SAMPLE.xlsx", "Agency Bonus.xlsx", "Agency Charges.xlsx")
    headings = Array("Commissions", "Bonuses", "Charges")

    ' Check if the "AgencyBenefit" sheet exists, if not, create it
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("AgencyBenefit")
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create a new sheet if "AgencyBenefit" does not exist
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "AgencyBenefit"
    End If

    ' Clear any existing data in the sheet
    ws.Cells.Clear

    ' Remove gridlines from the sheet
    ws.Application.ActiveWindow.DisplayGridlines = False

    ' Start copying data from row 20
    currentRow = 20

    ' Loop through all files
    For i = LBound(fileNames) To UBound(fileNames)
        ' Check if the file exists
        If Dir(excelFilePath & fileNames(i)) = "" Then
            MsgBox "File not found: " & excelFilePath & fileNames(i), vbExclamation
            Exit Sub
        End If

        ' Open the Excel file
        Set importedWorkbook = Workbooks.Open(excelFilePath & fileNames(i))

        ' Ensure the first sheet has data to copy
        If Application.WorksheetFunction.CountA(importedWorkbook.Sheets(1).UsedRange) > 0 Then
            ' Insert the heading for the current dataset
            ws.Cells(currentRow, 1).Value = headings(i)
            ws.Cells(currentRow, 1).Font.Bold = True
            With ws.Rows(currentRow)
                .HorizontalAlignment = xlLeft ' Left-align for headings (Commissions, Bonuses, Charges)
            End With
            ApplyBorders ws, currentRow, 7, True ' Apply borders and fill
            currentRow = currentRow + 1

            ' Remove the first blank row directly following the heading row
            If Application.WorksheetFunction.CountA(ws.Rows(currentRow)) = 0 Then
                ws.Rows(currentRow).Delete
            End If

            ' Copy dataset
            importedWorkbook.Sheets(1).UsedRange.Copy
            ws.Cells(currentRow, 1).PasteSpecial Paste:=xlPasteValues

            ' Apply top and bottom borders to the header row
            lastCol = ws.Cells(currentRow, ws.Columns.Count).End(xlToLeft).Column
            With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Interior.Color = xlNone
                .Font.Bold = True
                .Font.Size = 12
            End With

            ' Center-align all columns
            ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow + importedWorkbook.Sheets(1).UsedRange.Rows.Count, lastCol)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow + importedWorkbook.Sheets(1).UsedRange.Rows.Count, lastCol)).VerticalAlignment = xlCenter

            ' Format "Compensation Formula" column as percentages
            Dim compCol As Integer
            compCol = Application.Match("Compensation Formula", ws.Rows(currentRow), 0)
            If Not IsError(compCol) Then
                ws.Range(ws.Cells(currentRow + 1, compCol), ws.Cells(ws.Rows.Count, compCol).End(xlUp)).NumberFormat = "0.00%"
            End If

            ' Update the current row pointer
            currentRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1

            ' Perform calculations for specific files
            If fileNames(i) = "Agency Commissions TEST SAMPLE.xlsx" Then
                totalCommissions = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(currentRow - importedWorkbook.Sheets(1).UsedRange.Rows.Count + 1, lastCol), ws.Cells(currentRow - 1, lastCol)))
                ws.Cells(currentRow, 1).Value = "Total Commissions"
                ws.Cells(currentRow, 1).Font.Bold = True
                ws.Cells(currentRow, lastCol).Value = totalCommissions
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ws.Cells(currentRow, lastCol).Font.Bold = True
                ApplyBorders ws, currentRow, lastCol, False ' Apply borders only, no fill
                currentRow = currentRow + 2 ' Leave a blank row
            End If

            If fileNames(i) = "Agency Bonus.xlsx" Then
                totalBonus = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(currentRow - importedWorkbook.Sheets(1).UsedRange.Rows.Count + 1, lastCol), ws.Cells(currentRow - 1, lastCol)))
                ws.Cells(currentRow, 1).Value = "Total Bonus"
                ws.Cells(currentRow, 1).Font.Bold = True
                ws.Cells(currentRow, lastCol).Value = totalBonus
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ws.Cells(currentRow, lastCol).Font.Bold = True
                ApplyBorders ws, currentRow, lastCol, False ' Apply borders only, no fill
                currentRow = currentRow + 2 ' Leave a blank row

                ' Add Gross Pay
                ws.Cells(currentRow, 1).Value = "Gross Pay"
                With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))
                    .Font.Bold = True
                    .Interior.Color = RGB(0, 0, 0) ' Black fill
                    .Font.Color = RGB(255, 255, 255) ' White font
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                ws.Cells(currentRow, lastCol).Value = totalCommissions + totalBonus
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ApplyBorders ws, currentRow, lastCol, False
                currentRow = currentRow + 2 ' Leave a blank row
            End If

            If fileNames(i) = "Agency Charges.xlsx" Then
                totalChargebacks = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(currentRow - importedWorkbook.Sheets(1).UsedRange.Rows.Count + 1, lastCol), ws.Cells(currentRow - 1, lastCol)))
                ws.Cells(currentRow, 1).Value = "Total Chargebacks"
                ws.Cells(currentRow, 1).Font.Bold = True
                ws.Cells(currentRow, 1).Font.Color = vbRed
                ws.Cells(currentRow, lastCol).Value = totalChargebacks
                ws.Cells(currentRow, lastCol).Font.Color = vbRed
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ws.Cells(currentRow, lastCol).Font.Bold = True
                ApplyBorders ws, currentRow, lastCol, False
                currentRow = currentRow + 2 ' Leave a blank row

                ' Add Deductions
                deductions = 0 ' Placeholder value
                ws.Cells(currentRow, 1).Value = "Deductions"
                ws.Cells(currentRow, 1).Font.Bold = True
                With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))
                    .HorizontalAlignment = xlCenter ' Center-align deductions row
                End With
                ws.Cells(currentRow, lastCol).Value = deductions
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ws.Cells(currentRow, lastCol).Font.Bold = True
                ApplyBorders ws, currentRow, lastCol, True
                currentRow = currentRow + 2 ' Leave a blank row

                ' Calculate Total Deductions
                totalDeductions = totalChargebacks + deductions
                ws.Cells(currentRow, 1).Value = "Total Deductions"
                With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))
                    .Font.Bold = True
                    .Interior.Color = RGB(105, 105, 105) ' Dark gray fill
                    .Font.Color = RGB(255, 255, 255) ' White font
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                ws.Cells(currentRow, lastCol).Value = totalDeductions
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ApplyBorders ws, currentRow, lastCol, False
                currentRow = currentRow + 2 ' Leave a blank row

                ' Calculate Net Pay
                netPay = (totalCommissions + totalBonus) - totalDeductions
                ws.Cells(currentRow, 1).Value = "Net Pay"
                With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))
                    .Font.Bold = True
                    .Interior.Color = RGB(0, 0, 0) ' Black fill
                    .Font.Color = RGB(255, 255, 255) ' White font
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                ws.Cells(currentRow, lastCol).Value = netPay
                ws.Cells(currentRow, lastCol).NumberFormat = "$#,##0.00"
                ApplyBorders ws, currentRow, lastCol, False
                currentRow = currentRow + 2 ' Leave a blank row

                ' Add Year To Date
                ws.Cells(currentRow, 1).Value = "Year To Date"
                ws.Cells(currentRow, 1).Font.Bold = True
                With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 2))
                    .Interior.Color = RGB(211, 211, 211) ' Light gray fill
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
                ws.Cells(currentRow, 2).Value = 0
                ws.Cells(currentRow, 2).Font.Bold = True
                ApplyBorders ws, currentRow, 2, False
                currentRow = currentRow + 2 ' Leave a blank row
            End If
        End If
        importedWorkbook.Close SaveChanges:=False
    Next i

    ' Notify completion
    MsgBox "Data from all files imported successfully into the 'AgencyBenefit' sheet with formatted layout!", vbInformation
End Sub

Sub ApplyBorders(ws As Worksheet, rowIndex As Long, lastCol As Integer, includeFill As Boolean)
    ' Apply top and bottom borders to the specified row
    With ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        If includeFill Then
            .Interior.Color = RGB(211, 211, 211) ' Light gray fill
        End If
    End With
End Sub



2. Creating History/Record if data add or delete from the workspace

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsHistory As Worksheet
    Dim wsCurrent As Worksheet
    Dim lastRow As Long
    Dim i As Integer
    
    ' Set the worksheets
    Set wsCurrent = ThisWorkbook.Sheets("AgencyBenefit")
    Set wsHistory = ThisWorkbook.Sheets("Histor")
    
    ' Check if the change is in the "MixedFruit" sheet
    If Not Intersect(Target, wsCurrent.UsedRange) Is Nothing Then
        ' Find the last row in the History sheet
        lastRow = wsHistory.Cells(wsHistory.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Add the timestamp
        wsHistory.Cells(lastRow, 1).Value = Now
        
        ' Copy the changed row to the History sheet
        For i = 1 To wsCurrent.UsedRange.Columns.Count
            wsHistory.Cells(lastRow, i + 1).Value = wsCurrent.Cells(Target.Row, i).Value
        Next i
    End If
End Sub
