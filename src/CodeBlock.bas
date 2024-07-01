Attribute VB_Name = "CodeBlock"
Global oFSO As Object
Global oShell As Object

' Subroutine to initialize the application settings
Sub initializer()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

' Subroutine to restore the application settings
Sub destroyer()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' Subroutine to initialize the summary sheet
Sub initSummary()
    Dim i, n As Long
    initializer
    SUMMARY.Cells.Clear ' Clear the summary sheet
    n = MAPPER.Range("Map").Rows.Count ' Get the number of rows in the "Map" range in MAPPER sheet

    ' Copy "Variable" column from MAPPER to SUMMARY and set header values
    MAPPER.Range("Map[Variable]").Copy
    SUMMARY.Range("A1").Value = "ID"
    SUMMARY.Range("B1").Value = "File Name"
    SUMMARY.Range("C1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    ' Copy "ID" column from MAPPER to SUMMARY and set initial values
    MAPPER.Range("Map[ID]").Copy
    SUMMARY.Range("A2").Value = 0
    SUMMARY.Range("B2").Value = "."
    SUMMARY.Range("C2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    ' Create a table in SUMMARY and set alignment for headers
    SUMMARY.ListObjects.Add(xlSrcRange, SUMMARY.Range(SUMMARY.Cells(2, 1), SUMMARY.Cells(3, n + 2)), , xlYes).Name = "Summ"
    SUMMARY.Range("Summ[#Headers]").HorizontalAlignment = xlCenter
    SUMMARY.Range("Summ[#Headers]").VerticalAlignment = xlCenter
    
    ' Add formula to "A3" and apply styles and number formats based on "Type"
    SUMMARY.Range("A3").Formula = "=IF(LEN([@[.]])<>0,ROW()-ROW(Summ[[#Headers],[0]]),"""")"
    For i = 1 To n Step 1
        If MAPPER.Range("Map[Type]")(i) = "Output" Then
            SUMMARY.Range("Summ[" & i & "]").Style = "Output"
            SUMMARY.Range("Summ[" & i & "]").NumberFormat = "0.000"
        Else
            SUMMARY.Range("Summ[" & i & "]").Style = "Input"
        End If
    Next
    
    ' Auto-fit columns in the "Summ" table
    SUMMARY.Range("Summ").Columns.EntireColumn.AutoFit
    destroyer
End Sub

' Subroutine to get the template Excel object
Sub getTemplateExcelObject(ByRef wb As Workbook)
    Set wb = Nothing
    On Error Resume Next ' Continue running the macro even if an error occurs
    Set oFSO = CreateObject("Scripting.FileSystemObject") ' Create FileSystemObject
    Set wb = Workbooks(oFSO.GetFileName(SETTINGS.Range("InputTemplate").text)) ' Try to set workbook if already open
    If Err.Number <> 0 Then ' If error occurs (workbook not open), open the workbook
        Set wb = Workbooks.Open(SETTINGS.Range("InputTemplate").text, False, False)
        Err.Number = 0
    End If
End Sub

' Subroutine to generate data for one entry
Sub genOne()
    Dim wb As Workbook
    Dim i, j, m, n As Long
    Dim varsht, varrng As String
    getTemplateExcelObject wb ' Get the template workbook
    m = SUMMARY.Range("Summ").Rows.Count ' Get the number of rows in the "Summ" table
    n = MAPPER.Range("Map").Rows.Count ' Get the number of rows in the "Map" range
    i = SETTINGS.Range("SelectedItem").Value ' Get the selected item index

    If i <= m Then ' Check if the selected item index is within the range
        For j = 1 To n Step 1
            varsht = MAPPER.Range("Map[Sheet]")(j) ' Get the sheet name
            varrng = MAPPER.Range("Map[Reference]")(j) ' Get the cell reference
            If MAPPER.Range("Map[Type]")(j) = "Output" Then
                ' Copy value from template workbook to SUMMARY if type is "Output"
                SUMMARY.Range("Summ[" & j & "]")(i).Value = wb.Sheets(varsht).Range(varrng).Value
            ElseIf MAPPER.Range("Map[TYPE]")(j) = "Input" Then
                ' Copy value from SUMMARY to template workbook if type is "Input"
                wb.Sheets(varsht).Range(varrng).Value = SUMMARY.Range("Summ[" & j & "]")(i).Value
            End If
        Next
    End If
End Sub

' Subroutine to print or save one entry
Sub printOne()
    Dim wb As Workbook
    Dim i, k As Long
    Dim varsht, varrng, fpath, fname, extn As String
    Dim fileformatstr As Integer
    
    i = SETTINGS.Range("ExportIndex").Value ' Get the export index
    extn = SETTINGS.Range("extn").Cells(i, 1).Value ' Get the file extension
    fileformatstr = SETTINGS.Range("extn").Cells(i, 2).Value ' Get the file format string
    i = SETTINGS.Range("SelectedItem").Value ' Get the selected item index
    
    fpath = SETTINGS.Range("OutputFolder").Value ' Get the output folder path
    fpath = fpath & IIf(Right(fpath, 1) = "\", "", "\") ' Ensure the folder path ends with a backslash
    
    genOne ' Generate data for the selected item
    getTemplateExcelObject wb ' Get the template workbook
    fname = IIf(Len(SUMMARY.Range("Summ[.]")(i).Value) = 0, "temp", SUMMARY.Range("Summ[.]")(i).Value) & extn ' Set the file name
    If extn = ".pdf" Then
        ' Export as PDF if extension is ".pdf"
        wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fpath & fname, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False
    Else
        ' Save as specified format
        wb.SaveAs fpath & fname, fileformat:=fileformatstr
        wb.Close
    End If
End Sub

' Subroutine to open Excel files in the specified folder
Sub openExcelFiles()
    
    Dim folder As String, file As String
    Dim maprng As Range
    Set maprng = MAPPER.Range("Map") ' Set the mapping range
    folder = SETTINGS.Range("OutputFolder").Value ' Get the output folder path
    
    file = Dir(folder & "\") ' Get the first file in the folder
    While (file <> "") ' Loop through all files in the folder
        If InStr(LCase(Right(file, 5)), ".xls") > 0 Then
            ProcessSpreadsheet folder, file, maprng ' Process the spreadsheet if it is an Excel file
        End If
        file = Dir ' Get the next file
    Wend
End Sub

' Subroutine to process individual spreadsheets
Sub ProcessSpreadsheet(ByVal folder As String, ByVal file As String, ByRef maprng As Range)
    Dim wb As Workbook
    Dim rec As Long
    Dim i As Long
    On Error Resume Next ' Continue running the macro even if an error occurs
    rec = InsertRow(SUMMARY, "Summ") ' Insert a new row in the summary table
    ' Add a hyperlink to the file in the summary table
    SUMMARY.Hyperlinks.Add Anchor:=SUMMARY.Range("Summ[.]").Cells(rec, 1), _
            Address:=folder & "\" & file, ScreenTip:=file, TextToDisplay:=file
    Set wb = Workbooks.Open(folder & "\" & file, False, True) ' Open the workbook
    For i = 1 To maprng.Rows.Count
        ' Copy values from the workbook to the summary table based on the mapping range
        SUMMARY.Range("Summ[" & maprng.Cells(i, 1).text & "]").Cells(rec, 1).Value = _
                wb.Worksheets(maprng.Cells(i, 3).text).Range(maprng.Cells(i, 4).text).text
    Next i
    wb.Close SaveChanges:=False ' Close the workbook without saving changes
End Sub

' Function to insert a new row in the specified table
Function InsertRow(sht As Worksheet, tabnme As String) As Variant
    If Len(sht.Range(tabnme).Cells(sht.Range(tabnme).Rows.Count, 1).Value) <> 0 Then
        ' Add a new row to the table if the last row is not empty
        sht.Range(tabnme).ListObject.ListRows.Add AlwaysInsert:=True
    End If
    InsertRow = sht.Range(tabnme).Rows.Count ' Return the new row count
End Function



