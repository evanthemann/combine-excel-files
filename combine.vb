Sub CombineExcelFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim wbDest As Workbook
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim destFileName As String
    Dim lastSheetIndex As Integer
    Dim sheetIndex As Integer

    ' Prompt the user to enter the folder path
    folderPath = InputBox("Please enter the folder path where the Excel files are located:") & "\"
    
    ' Determine the destination file name
    destFileName = folderPath & "combined.xlsx"
    
    ' Create a new workbook for combining sheets
    Set wbDest = Workbooks.Add
    wbDest.SaveAs fileName:=destFileName
    Set wbDest = Workbooks.Open(destFileName)
    
    ' Initialize the file search
    fileName = Dir(folderPath & "*.xlsx")
    sheetIndex = 0  ' Initialize sheet index
    
    ' Loop through all files in the folder
    Do While fileName <> ""
        ' Check that the file is not the destination file
        If fileName <> "combined.xlsx" Then
            ' Open the source workbook
            On Error Resume Next
            Set wbSource = Workbooks.Open(folderPath & fileName)
            If Err.Number <> 0 Then
                MsgBox "Could not open file: " & folderPath & fileName
                Err.Clear
                GoTo NextFile
            End If
            On Error GoTo 0
            
            ' Loop through all sheets in the source workbook
            For Each wsSource In wbSource.Sheets
                ' Check if the sheet is visible before copying
                If wsSource.Visible = xlSheetVisible Then
                    ' Attempt to copy the sheet to the destination workbook
                    On Error Resume Next
                    wsSource.Copy After:=wbDest.Sheets(wbDest.Sheets.Count)
                    If Err.Number <> 0 Then
                        MsgBox "Error copying sheet: " & wsSource.Name & " from " & fileName
                        Err.Clear
                    Else
                        ' Rename the sheet to avoid conflicts
                        sheetIndex = sheetIndex + 1
                        wbDest.Sheets(wbDest.Sheets.Count).Name = "Sheet" & sheetIndex
                    End If
                    On Error GoTo 0
                End If
            Next wsSource
            
            ' Close the source workbook without saving changes
            wbSource.Close SaveChanges:=False
        End If
    
NextFile:
        ' Get the next file
        fileName = Dir
    Loop

    ' Optional: Delete any default sheets in the destination workbook if necessary
    Application.DisplayAlerts = False
    For lastSheetIndex = wbDest.Sheets.Count To 2 Step -1
        If wbDest.Sheets(lastSheetIndex).Name Like "Sheet*" Then
            If IsEmpty(wbDest.Sheets(lastSheetIndex).UsedRange) Then
                wbDest.Sheets(lastSheetIndex).Delete
            End If
        End If
    Next lastSheetIndex
    Application.DisplayAlerts = True
    
    ' Save and close the combined workbook
    wbDest.Save
    wbDest.Close
    
    ' Inform the user that the operation is complete
    MsgBox "All sheets have been combined into " & destFileName

End Sub

