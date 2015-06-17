Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form1

    Private Sub btnMerge_Click(sender As Object, e As EventArgs) Handles btnMerge.Click
        If (txtCircleLeaderReportLocation.Text.Trim.Length > 0) And (txtCIReportFileLocation.Text.Trim.Length > 0) Then
            Dim CircleLeaderFile As String = txtCircleLeaderReportLocation.Text.Trim
            Dim CIReportFile As String = txtCIReportFileLocation.Text.Trim

            If File.Exists(CircleLeaderFile) And File.Exists(CIReportFile) Then

                ' CircleLeader Excel Information
                Dim CircleLeaderApp As New Excel.Application
                Dim CircleLeaderWorksheet As Excel.Worksheet
                Dim CircleLeaderWorkbook As Excel.Workbook
                CircleLeaderWorkbook = CircleLeaderApp.Workbooks.Open(CircleLeaderFile)
                CircleLeaderWorksheet = CircleLeaderWorkbook.Sheets.Item(1)

                ' To read all columns from the Circle Leader and load it to an Array
                Dim CircleLeaderUsedRange As Excel.Range = CircleLeaderWorksheet.UsedRange
                Dim CircleLeaderArray As System.Array = CircleLeaderUsedRange.Value

                ' CIReport Excel Information
                Dim CIApp As New Excel.Application
                Dim CIWorksheet As Excel.Worksheet
                Dim CIWorkbook As Excel.Workbook
                CIWorkbook = CIApp.Workbooks.Open(CIReportFile)
                CIWorksheet = CIWorkbook.Sheets.Item(1)

                ' To read all columns from the CI Report and load it to an Array
                Dim CIReportUsedRange As Excel.Range = CIWorksheet.UsedRange
                Dim CIReportArray As System.Array = CIReportUsedRange.Value

                ' Final Merged File
                Dim MergedFileApp As New Excel.Application
                Dim MergedFileWorkBook As Excel.Workbook
                Dim MergedFileWorksheet As Excel.Worksheet
                Dim MergedFile As String
                MergedFile = txtMergedFileLocation.Text.Trim + "\" + txtMergedFileName.Text.Trim
                MergedFileWorkBook = MergedFileApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
                MergedFileWorksheet = MergedFileWorkBook.Sheets.Item(1)

                Dim NewRow As Integer = 1  ' Keeps track of the Rows in the Merged File
                Dim NewCol As Integer = 1  ' Keeps track of the cols in the Merged File

                Dim firstRow As Boolean = False
                Dim currentRow As Integer
                Dim val As String
                For row As Integer = 3 To (CircleLeaderArray.GetUpperBound(0) - 2)
                    NewCol = 1
                    For col As Integer = 1 To (CircleLeaderArray.GetUpperBound(1))
                        Select Case col
                            Case 1, 2, 3, 4, 5, 7, 8, 9, 14, 22, 23, 24, 30
                                MergedFileWorksheet.Cells(NewRow, NewCol).Value = CircleLeaderArray(row, col)
                                NewCol = NewCol + 1
                        End Select
                    Next
                    firstRow = True
                    currentRow = NewRow

                    If currentRow = 1 Then
                        'This is to write the Header 
                        For col As Integer = 3 To 66  ' (CIReportArray.GetUpperBound(1) - 3)
                            Select Case col
                                Case 3, 4, 7, 20, 21, 22, 24, 25, 26, 27, 28, 29, 30, 35, 41, 42, 47, 48, 66
                                    MergedFileWorksheet.Cells(NewRow, NewCol).NumberFormat = "@" ' Change format of cell to store number as text
                                    MergedFileWorksheet.Cells(NewRow, NewCol).Value = CIReportArray(2, col)
                                    NewCol = NewCol + 1
                            End Select
                        Next
                    Else
                        For CIrow As Integer = 2 To (CIReportArray.GetUpperBound(0) - 2)
                            If (CIReportArray(CIrow, 2).Equals(CircleLeaderArray(row, 1))) Then
                                If firstRow = False Then
                                    '' If it the first row for CIReport then you dont need to populate the 
                                    '' the rest of the information belonging to Circle Leader Report in the initial column
                                    '' as it is already populated
                                    NewCol = 1
                                    NewRow = NewRow + 1
                                    For col As Integer = 1 To (CircleLeaderArray.GetUpperBound(1))
                                        Select Case col
                                            Case 1, 2, 3, 4, 5, 7, 8, 9, 14, 22, 23, 24, 30
                                                MergedFileWorksheet.Cells(NewRow, NewCol).Value = CircleLeaderArray(row, col)
                                                NewCol = NewCol + 1
                                        End Select
                                    Next

                                Else
                                    'startRow = NewRow
                                    firstRow = False
                                End If
                                For col As Integer = 3 To 66  ' (CIReportArray.GetUpperBound(1) - 3)
                                    Select Case col
                                        Case 4
                                            Select Case CIReportArray(CIrow, col)
                                                Case "Graduation Report"
                                                    val = "01 Graduation Report"
                                                Case "Initial Report"
                                                    val = "00 Initial Report"
                                                Case "6 Month Report"
                                                    val = "06 Month Report"
                                                Case Else
                                                    val = CIReportArray(CIrow, col)
                                            End Select
                                            MergedFileWorksheet.Cells(NewRow, NewCol).NumberFormat = "@" ' Change format of cell to store number as text
                                            MergedFileWorksheet.Cells(NewRow, NewCol).Value = val
                                            NewCol = NewCol + 1
                                        Case 3, 7, 20, 21, 22, 24, 25, 26, 27, 28, 29, 30, 35, 41, 42, 47, 48, 66
                                            MergedFileWorksheet.Cells(NewRow, NewCol).NumberFormat = "@" ' Change format of cell to store number as text
                                            MergedFileWorksheet.Cells(NewRow, NewCol).Value = CIReportArray(CIrow, col)
                                            NewCol = NewCol + 1
                                    End Select
                                Next
                            End If
                        Next
                    End If
                    'endRow = NewRow
                    'If startRow = endRow Or startRow = 0 Then
                    '' No Merge is needed
                    'Else
                    'For column As Integer = 1 To 13
                    'MergedFileWorksheet.Range(MergedFileWorksheet.Cells(startRow, column), MergedFileWorksheet.Cells(endRow, column)).Merge()
                    'Next
                    'End If
                    NewRow = NewRow + 1
                Next

                ' sort the Merged Files
                Dim MergedRange As Excel.Range = MergedFileWorksheet.Application.Range("A1", "AF" + Convert.ToString(NewRow - 1))
                MergedRange.Sort(Key1:=MergedRange.Columns(1), Order1:=Excel.XlSortOrder.xlAscending, _
                                 Key2:=MergedRange.Columns(15), Order2:=Excel.XlSortOrder.xlAscending, _
                                 Orientation:=Excel.XlSortOrientation.xlSortColumns, _
                                Header:=Excel.XlYesNoGuess.xlYes, _
                                SortMethod:=Excel.XlSortMethod.xlPinYin, _
                                DataOption1:=Excel.XlSortDataOption.xlSortNormal, _
                                DataOption2:=Excel.XlSortDataOption.xlSortNormal)

                ' Now that it is sorted you can remove the duplicate ID 
                Dim previousValue As String = "0"
                Dim currentValue As String
                For row As Integer = 2 To (NewRow - 1)
                    currentValue = MergedFileWorksheet.Cells(row, 1).value.ToString
                    If currentValue <> previousValue Then
                        previousValue = currentValue
                    Else
                        For column As Integer = 1 To 13
                            MergedFileWorksheet.Cells(row, column).value = ""
                        Next
                    End If

                Next

                'Merge the cells
                Dim startRow As Integer = 0
                Dim endRow As Integer = 0
                previousValue = "0"
                currentValue = "0"
                For row As Integer = 2 To (NewRow - 1)

                    If MergedFileWorksheet.Cells(row, 1).value Is Nothing Then
                        endRow = row
                    Else
                        currentValue = MergedFileWorksheet.Cells(row, 1).value.ToString
                        If startRow < endRow And startRow <> 0 Then
                            For column As Integer = 1 To 13
                                MergedFileWorksheet.Range(MergedFileWorksheet.Cells(startRow, column), MergedFileWorksheet.Cells(endRow, column)).Merge()
                            Next
                        End If
                        startRow = row
                    End If
                Next
                'Close the files
                MergedFileWorkBook.Close(SaveChanges:=True, Filename:=MergedFile)
                MergedFileApp.Quit()
                CircleLeaderApp.ActiveWorkbook.Close(False, CircleLeaderFile)
                CircleLeaderApp.Quit()
                CIApp.ActiveWorkbook.Close(False, CircleLeaderFile)
                CIApp.Quit()
                MsgBox("File Created at :" + MergedFile)
            End If
        End If
    End Sub


End Class
