Imports Microsoft.Win32
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices
Class MainWindow
    Dim selectedImportFile As String = ""
    Dim selectedFiles As String()

    Private Sub ChooseFileBtn_Click(sender As Object, e As RoutedEventArgs) Handles ChooseFileBtn.Click
        Dim openFileDialog As New OpenFileDialog() ' Create an instance of OpenFileDialog

        ' Set filter options and filter index
        openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx"
        openFileDialog.FilterIndex = 1
        openFileDialog.Multiselect = True ' Allow multiple file selection

        Dim result As Boolean? = openFileDialog.ShowDialog() ' Call the ShowDialog method to show the dialog box

        If result = True Then ' Process input if the user clicked OK
            selectedFiles = openFileDialog.FileNames ' Retrieve the selected file paths

            If selectedFiles.Length > 0 Then
                ' Update button content with the name of the first selected file
                ChooseFileBtn.Content = System.IO.Path.GetFileName(selectedFiles(0))
            Else
                MsgBox("No files selected", MsgBoxStyle.Information, "Select Files")
            End If
        End If
    End Sub

    Private Async Sub ExtractDataBtn_Click(sender As Object, e As RoutedEventArgs) Handles ExtractDataBtn.Click
        ' Ensure selectedFiles is properly populated with multiple file paths
        If selectedFiles Is Nothing OrElse selectedFiles.Length = 0 Then
            MsgBox("No files selected", MsgBoxStyle.Information, "Select Files")
            Return
        End If

        Dim activityTitle As New ListBoxItem()
        activityTitle.Content = "Registration data extract"
        activityTitle.Background = GetBrushFromHex("#3308E0FF")
        activityTitle.Padding = New Thickness(5, 2, 5, 2)
        activityTitle.Margin = New Thickness(0, 3, 0, 0)
        activityTitle.FontWeight = FontWeights.Bold
        activityTitle.BorderBrush = StatusColorHex("Default")
        activityTitle.BorderThickness = New Thickness(3, 0, 0, 0)
        activityTitle.Tag = (DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds
        ExtractedDataList.Items.Add(activityTitle) ' Add the styled item to the ListBox

        Await Task.Run(Sub()
                           For Each importExcelFile As String In selectedFiles
                               Dim excelApp As Excel.Application = Nothing
                               Dim workbook As Excel.Workbook = Nothing
                               Dim worksheet As Excel.Worksheet = Nothing
                               Dim newWorkbook As Excel.Workbook = Nothing
                               Dim newWorksheet As Excel.Worksheet = Nothing

                               Try
                                   excelApp = New Excel.Application()
                                   excelApp.Visible = False ' Make Excel visible (optional).
                                   workbook = excelApp.Workbooks.Open(importExcelFile)
                                   worksheet = workbook.Sheets("Sheet1")

                                   ' Create a new workbook and worksheet
                                   newWorkbook = excelApp.Workbooks.Add()
                                   newWorksheet = newWorkbook.Sheets(1)

                                   ' Create headers (same as before, can be refactored into a separate method if needed)
                                   newWorksheet.Range("A1").Value = "AUTORGS: ELEMENTAL IMPORT TEMPLATE"
                                   newWorksheet.Range("A1:E1").Merge()
                                   newWorksheet.Range("A1:E1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                   newWorksheet.Range("A1:E1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                   newWorksheet.Range("A1:E1").Font.Size = 20
                                   newWorksheet.Range("A1:E1").Font.Bold = True

                                   newWorksheet.Range("A2").Value = "EXAMINATION BOARD"
                                   newWorksheet.Range("A2:E2").Merge()
                                   newWorksheet.Range("A2:E2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                   newWorksheet.Range("A2:E2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                   newWorksheet.Range("A2:E2").Font.Size = 16
                                   newWorksheet.Range("A2:E2").Font.Bold = False

                                   newWorksheet.Range("A3").Value = "CANDIDATE MARKS"
                                   newWorksheet.Range("A3:E3").Merge()
                                   newWorksheet.Range("A3:E3").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                   newWorksheet.Range("A3:E3").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                   newWorksheet.Range("A3:E3").Font.Size = 16
                                   newWorksheet.Range("A3:E3").Font.Bold = True

                                   newWorksheet.Range("A4").Value = "CANDIDATE DETAILS"
                                   newWorksheet.Range("A4:E4").Merge()
                                   newWorksheet.Range("A4:E4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                   newWorksheet.Range("A4:E4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                   newWorksheet.Range("A4:E4").Font.Size = 11
                                   newWorksheet.Range("A4:E4").Font.Bold = True
                                   newWorksheet.Range("A4:E4").Interior.Color = System.Drawing.Color.LightGreen ' Fill Light Green

                                   newWorksheet.Range("A5").ColumnWidth = 4
                                   newWorksheet.Range("B5").ColumnWidth = 10
                                   newWorksheet.Range("C5").ColumnWidth = 9
                                   newWorksheet.Range("D5").ColumnWidth = 35
                                   newWorksheet.Range("E5").ColumnWidth = 8

                                   newWorksheet.Range("A5").Value = "#"
                                   newWorksheet.Range("B5").Value = "SCH. CODE"
                                   newWorksheet.Range("C5").Value = "CAND. NO."
                                   newWorksheet.Range("D5").Value = "NAME"
                                   newWorksheet.Range("E5").Value = "GENDER"

                                   With newWorksheet.Range("A4:E5").Borders
                                       .LineStyle = Excel.XlLineStyle.xlContinuous
                                       .ColorIndex = Excel.XlRgbColor.rgbBlack
                                       .TintAndShade = 0
                                       .Weight = Excel.XlBorderWeight.xlThin
                                   End With

                                   newWorksheet.Range("A5:E5").Font.Size = 11
                                   newWorksheet.Range("A5:E5").Font.Bold = True
                                   newWorksheet.Range("A5:E5").Interior.Color = System.Drawing.ColorTranslator.FromHtml("#8DBE5C") ' Fill Green Accent 6 Lighter 40%

                                   ' Copy data from the original worksheet to the new worksheet
                                   Dim indexRow As Integer = 6
                                   Dim sourceIndexRow As Integer = 2
                                   Dim newRowIndex As Integer = 1
                                   Do While worksheet.Range("A" & indexRow).Value IsNot Nothing
                                       Dim myString As String = worksheet.Range("C" & sourceIndexRow).Value
                                       Dim parts() As String = myString.Split("/"c)
                                       Dim result As String = If(parts.Length > 1, parts(1), String.Empty)

                                       newWorksheet.Range("A" & indexRow).Value = newRowIndex
                                       newWorksheet.Range("B" & indexRow & ":C" & indexRow).NumberFormat = "@"
                                       newWorksheet.Range("B" & indexRow).Value = worksheet.Range("B" & sourceIndexRow).Value
                                       newWorksheet.Range("C" & indexRow).Value = result
                                       newWorksheet.Range("D" & indexRow).Value = worksheet.Range("D" & sourceIndexRow).Value
                                       newWorksheet.Range("E" & indexRow).Value = worksheet.Range("F" & sourceIndexRow).Value

                                       With newWorksheet.Range("A" & indexRow & ":E" & indexRow).Borders
                                           .LineStyle = Excel.XlLineStyle.xlContinuous
                                           .ColorIndex = Excel.XlRgbColor.rgbBlack
                                           .TintAndShade = 0
                                           .Weight = Excel.XlBorderWeight.xlThin
                                       End With

                                       indexRow += 1
                                       newRowIndex += 1
                                       sourceIndexRow += 1
                                   Loop

                                   ' Save the new workbook
                                   Dim newFilePath As String = IO.Path.Combine(IO.Path.GetDirectoryName(importExcelFile), "Reprocessed ~ " & IO.Path.GetFileName(importExcelFile))
                                   newWorkbook.SaveAs(newFilePath)

                                   Dispatcher.Invoke(Sub()
                                                         Dim newSessionLog As New ListBoxItem()
                                                         newSessionLog.Content = "Ready: " & newFilePath
                                                         newSessionLog.Foreground = StatusColorHex("Info")
                                                         newSessionLog.Tag = (DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds
                                                         ExtractedDataList.Items.Add(newSessionLog)
                                                     End Sub)
                               Catch ex As Exception
                                   Dispatcher.Invoke(Sub()
                                                         Dim newSessionLog As New ListBoxItem()
                                                         newSessionLog.Content = "Error: " & importExcelFile & ": " & ex.Message
                                                         newSessionLog.Foreground = StatusColorHex("Info")
                                                         newSessionLog.Tag = (DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds
                                                         ExtractedDataList.Items.Add(newSessionLog)
                                                     End Sub)
                               Finally
                                   ' Clean up
                                   If newWorksheet IsNot Nothing Then Marshal.ReleaseComObject(newWorksheet)
                                   If newWorkbook IsNot Nothing Then
                                       newWorkbook.Close(False)
                                       Marshal.ReleaseComObject(newWorkbook)
                                   End If
                                   If worksheet IsNot Nothing Then Marshal.ReleaseComObject(worksheet)
                                   If workbook IsNot Nothing Then
                                       workbook.Close(False)
                                       Marshal.ReleaseComObject(workbook)
                                   End If
                                   If excelApp IsNot Nothing Then
                                       excelApp.Quit()
                                       Marshal.ReleaseComObject(excelApp)
                                   End If
                               End Try
                           Next
                       End Sub)
    End Sub

    Private Sub CopyAllContentToFile_Click(sender As Object, e As RoutedEventArgs) Handles CopyAllContentToFile.Click
        ' Check if multiple files are selected
        If selectedFiles IsNot Nothing AndAlso selectedFiles.Length > 0 Then
            For Each sourceFilePath In selectedFiles
                ' Generate new file path with "Processed ~ " prefix
                Dim directory As String = Path.GetDirectoryName(sourceFilePath)
                Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(sourceFilePath)
                Dim fileExtension As String = Path.GetExtension(sourceFilePath)
                Dim destinationFilePath As String = Path.Combine(directory, "Processed ~ " & fileNameWithoutExtension & fileExtension)

                ProcessFile(sourceFilePath, destinationFilePath)
            Next
        ElseIf Not String.IsNullOrEmpty(selectedImportFile) Then
            ' Process a single file
            Dim sourceFilePath As String = selectedImportFile
            Dim directory As String = Path.GetDirectoryName(sourceFilePath)
            Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(sourceFilePath)
            Dim fileExtension As String = Path.GetExtension(sourceFilePath)
            Dim destinationFilePath As String = Path.Combine(directory, "Processed ~ " & fileNameWithoutExtension & fileExtension)

            ProcessFile(sourceFilePath, destinationFilePath)
        Else
            MsgBox("No file selected", MsgBoxStyle.Information, "Select File")
        End If
    End Sub

    Private Sub ProcessFile(sourceFilePath As String, destinationFilePath As String)
        Dim excelApp As New Excel.Application
        Dim sourceWorkbook As Excel.Workbook = Nothing
        Dim destinationWorkbook As Excel.Workbook = Nothing
        Dim sourceWorksheet As Excel.Worksheet = Nothing
        Dim destinationWorksheet As Excel.Worksheet = Nothing

        Try
            ' Open source workbook and worksheet
            sourceWorkbook = excelApp.Workbooks.Open(sourceFilePath)
            sourceWorksheet = sourceWorkbook.Sheets(1) ' Assuming data is in the first sheet

            ' Create a new workbook and worksheet
            destinationWorkbook = excelApp.Workbooks.Add()
            destinationWorksheet = destinationWorkbook.Sheets(1) ' Use the first sheet of the new workbook

            ' Copy all content from source to destination
            Dim usedRange As Excel.Range = sourceWorksheet.UsedRange
            usedRange.Copy()
            destinationWorksheet.Paste()

            ' Save the new workbook with prefix
            destinationWorkbook.SaveAs(destinationFilePath)
            Dim activityTitle As New ListBoxItem()
            activityTitle.Content = "Ready: " & destinationFilePath
            activityTitle.Background = GetBrushFromHex("#3308E0FF")
            activityTitle.Padding = New Thickness(5, 2, 5, 2)
            activityTitle.Margin = New Thickness(0, 3, 0, 0)
            activityTitle.FontWeight = FontWeights.Bold
            activityTitle.BorderBrush = StatusColorHex("Default")
            activityTitle.BorderThickness = New Thickness(3, 0, 0, 0)
            activityTitle.Tag = (DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds
            ExtractedDataList.Items.Add(activityTitle) ' Add the styled item to the ListBox

            'MsgBox("Data copied and file saved successfully as: " & destinationFilePath, MsgBoxStyle.Information, "Success")

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ' Clean up
            If sourceWorksheet IsNot Nothing Then Marshal.ReleaseComObject(sourceWorksheet)
            If sourceWorkbook IsNot Nothing Then
                sourceWorkbook.Close(False)
                Marshal.ReleaseComObject(sourceWorkbook)
            End If
            If destinationWorksheet IsNot Nothing Then Marshal.ReleaseComObject(destinationWorksheet)
            If destinationWorkbook IsNot Nothing Then
                destinationWorkbook.Close(False)
                Marshal.ReleaseComObject(destinationWorkbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    Private Sub ChooseMultiple_Click(sender As Object, e As RoutedEventArgs) Handles ChooseMultiple.Click
        Dim openFileDialog As New OpenFileDialog()

        ' Set filter options and filter index
        openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx"
        openFileDialog.FilterIndex = 1
        openFileDialog.Multiselect = True ' Allow multiple file selection

        Dim result As Boolean? = openFileDialog.ShowDialog() ' Call the ShowDialog method to show the dialog box

        If result = True Then ' Process input if the user clicked OK
            selectedFiles = openFileDialog.FileNames ' Get the selected file paths
            If selectedFiles.Length > 0 Then
                ' Optionally, update the UI with the selected file names
                Dim fileNames As String = String.Join(", ", selectedFiles.Select(Function(f) Path.GetFileName(f)))
                ChooseFileBtn.Content = fileNames ' Set the button content to show file names

                ' Process each selected file
                For Each filePath In selectedFiles
                    ' Handle each file here
                    ' For example, store them in a list or process them as needed
                    ' You could add code here to process the files
                Next
            End If
        End If
    End Sub
End Class
