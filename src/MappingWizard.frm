VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MappingWizard 
   Caption         =   "Synthesizer > Map Data Wizard"
   ClientHeight    =   4308
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8196.001
   OleObjectBlob   =   "MappingWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MappingWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wb As Workbook
Const colVariable As Integer = 2
Const colSheet As Integer = 3
Const colReference As Integer = 4
Const colType As Integer = 5


Private Sub AutomaticallyTriggerCellPicker_Click()
    SETTINGS.Range("AutomaticallyTriggerCellPicker").Value = Me.AutomaticallyTriggerCellPicker.Value
End Sub

Private Sub CancelButton_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub CellNameCheckBox_Click()
     SETTINGS.Range("UseCellNameInsteadOfAddress").Value = Me.CellNameCheckBox.Value
End Sub



Private Sub HelpButton_Click()
    ' Open the help link
    'Dim helpURL As String
    'helpURL = "https://www.example.com/help-document" ' Change this to your help URL
    'ThisWorkbook.FollowHyperlink Address:=helpURL, NewWindow:=True
End Sub

Private Sub MappedCellPickerButton_Click()
    ThisWorkbook.Activate
    getTemplateExcelObject wb
    wb.Activate

    Dim cellInfo As String
    Dim delimiterPosition As Integer
    
    ' Call the GetCellAddress function and store the result
    cellInfo = CellPicker("Please select a cell to be mapped that contains the value of the variable:", _
        "Mapping Value Cell Selection", Me.CellNameCheckBox.Value)
    ThisWorkbook.Activate
    ' Display the cell information
    If cellInfo <> "" Then
        delimiterPosition = InStr(1, cellInfo, "/", vbTextCompare)
        Me.SheetNamePicker = Left(cellInfo, delimiterPosition - 1)
        Me.CellReference = Right(cellInfo, Len(cellInfo) - delimiterPosition)
    End If
    
End Sub


Private Sub MappedVariablePickerButton_Click()
    ThisWorkbook.Activate
    getTemplateExcelObject wb
    wb.Activate
    
    If Me.SheetNamePicker.Value <> "" Then wb.Sheets(Me.SheetNamePicker.Value).Select
    If Me.CellReference.Value <> "" Then ActiveSheet.Range(Me.CellReference.Value).Select
    

    Dim cellInfo As String
    Dim delimiterPosition As Integer
    
    ' Call the GetCellAddress function and store the result
    cellInfo = CellPicker("Please select a cell whose value is Mapping Variable's label:", _
        "Variable Label Cell Selection", Me.CellNameCheckBox.Value)
    ' Display the cell information
    If cellInfo <> "" Then
        delimiterPosition = InStr(1, cellInfo, "/", vbTextCompare)
        Me.VariableName = ActiveSheet.Range(Right(cellInfo, Len(cellInfo) - delimiterPosition)).Value
    End If
    ThisWorkbook.Activate

End Sub

Private Function ValidateComboBox(ByRef cb As ComboBox) As Boolean
    Dim i As Integer
    
    ' Initialize the ValidateComboBox flag
    ValidateComboBox = False

    ' Loop through all items in the ComboBox
    For i = 1 To cb.ListCount
        If cb.Value = cb.List(i - 1) Then
            ValidateComboBox = True
            Exit For
        End If
    Next i

End Function

Private Function ValidateReference() As Boolean
    Dim reference As String
    On Error Resume Next
    reference = wb.Sheets(Me.SheetNamePicker.Value).Range(Me.CellReference.Value).Address
    ValidateReference = (Err.Number = 0)
    If Not ValidateReference Then Err.Number = 0
End Function

Private Function ValidateInputs() As Boolean
    Dim validSheet As Boolean, validReference As Boolean, validLabel As Boolean
    ValidateInputs = True
    
    ' Set the ComboBox to validate
    validSheet = ValidateComboBox(Me.SheetNamePicker)
    
    ' Check if the value is valid
    If Not validSheet Then
        messageBox "Please input a valid Template Sheet Name." & vbNewLine & Me.SheetNamePicker.Value, "Template Sheet Not Found", vbCritical
        Me.SheetNamePicker.Value = ""
        Me.SheetNamePicker.SetFocus
    End If
    
    ValidateInputs = ValidateInputs And validSheet
    If Not ValidateInputs Then Exit Function
    
    validReference = ValidateReference()
    
    If Not validReference Then
        messageBox "Please input a valid Template Cell Reference." & vbNewLine & Me.CellReference.Value, "Template Cell Reference Not Found", vbCritical
        Me.CellReference.Value = ""
        Me.CellReference.SetFocus
    End If
    
    ValidateInputs = ValidateInputs And validReference
    If Not ValidateInputs Then Exit Function
    
    validLabel = (Me.VariableName.Value <> "")
    
    If Not validLabel Then
        messageBox "Mapped Variable Name cannot be blank", "Missing Input", vbCritical
        Me.VariableName.SetFocus
    End If
    
    ValidateInputs = ValidateInputs And validLabel
    
End Function

Private Sub ResetButton_Click()
    CancelButton_Click
    MappingWizard.Show   ' Show the MappingWizard Form again
End Sub

Private Sub SaveAndNewButton_Click()
    SaveButton_Click
    MappingWizard.Show
End Sub

Private Sub SaveButton_Click()
    Dim valid As Boolean
    valid = ValidateInputs()
    InsertRecord
    CancelButton_Click
End Sub

Private Sub InsertRecord()
    Dim countRow As Long, rec As Long
    Dim sheetName As String, cellAddress As String, cellName As String, recAddress As String, selectedType As String
    Dim chosenCell As Range, mapperTable As Range
    Dim matchFound As Boolean
    Dim confirmReplace As VbMsgBoxResult
    Dim confirmMessage As String
    
    matchFound = False
    Set mapperTable = MAPPER.Range("Map")
    sheetName = Me.SheetNamePicker.Value
    cellName = ""
    cellAddress = ""
    Set chosenCell = wb.Sheets(sheetName).Range(Me.CellReference.Value)
    selectedType = IIf(Me.OutputOption, "Output", "Input")
    On Error Resume Next
    cellName = chosenCell.Name.Name
    cellAddress = UCase(Replace(chosenCell.Cells(1, 1).Address, "$", ""))
    On Error GoTo 0
    countRow = mapperTable.Rows.Count
    For rec = 1 To countRow Step 1
        If mapperTable.Cells(rec, colSheet).Value = sheetName Then
            recAddress = mapperTable.Cells(rec, colReference).Value
            If UCase(Replace(recAddress, "$", "")) = cellAddress Then
                matchFound = True
                Exit For
            ElseIf recAddress = cellName Then
                matchFound = True
                Exit For
            End If
        End If
    Next rec
    
    If matchFound Then
        confirmMessage = ""
        If mapperTable.Cells(rec, colVariable).Value <> Me.VariableName.Value Then
            confirmMessage = confirmMessage _
                & "Do you want to Rename the Variable from `" _
                & mapperTable.Cells(rec, colVariable).Value & "` to `" _
                & Me.VariableName.Value & "` ?"
        End If
        If mapperTable.Cells(rec, colType).Value <> selectedType Then
            If confirmMessage <> "" Then confirmMessage = confirmMessage & vbNewLine & vbNewLine
            confirmMessage = confirmMessage _
                & "Do you want to Chante the Variable Type from `" _
                & mapperTable.Cells(rec, colType).Value & "` to `" _
                & selectedType & "` ?"
        End If
        If confirmMessage = "" Then
            messageBox "Skipping Save since variable already mapped with same information in Table Row #" & rec, _
                "Save Skipped", vbInformation
        Else
            confirmMessage = "Found a mapping match in Table Row #" & rec & vbNewLine & vbNewLine & confirmMessage
            confirmReplace = messageBox(confirmMessage, "Replace Existing Mapping?", vbOKCancel)
            If confirmReplace = vbOK Then
                WriteRecord mapperTable, rec, selectedType
                messageBox "Updated Table Row #" & rec & " with the supplied mapping", _
                    "Matching Record Updated", vbInformation
            Else
                messageBox "Skipping Table Row #" & rec & " since you clicked cancel", _
                    "Skipped Record Update", vbInformation
            End If
        End If
    Else
        If mapperTable.Cells(countRow, colVariable).Value <> "" Then
            rec = countRow + 1
        Else
            rec = countRow
        End If
        WriteRecord mapperTable, rec, selectedType
    End If
End Sub

Sub WriteRecord(ByRef rng As Range, ByVal rec As Long, ByVal selectedType As String)
    rng.Cells(rec, colVariable).Value = Me.VariableName.Value
    rng.Cells(rec, colSheet).Value = Me.SheetNamePicker.Value
    rng.Cells(rec, colReference).Value = Me.CellReference.Value
    rng.Cells(rec, colType).Value = selectedType
End Sub

Private Sub UserForm_Initialize()
    Set wb = Nothing
    On Error Resume Next ' Continue running the macro even if an error occurs
    Set oFSO = CreateObject("Scripting.FileSystemObject") ' Create FileSystemObject
    Set wb = Workbooks(oFSO.GetFileName(SETTINGS.Range("InputTemplate").text)) ' Try to set workbook if already open
    If Err.Number <> 0 Then ' If error occurs (workbook not open), open the workbook
        Err.Number = 0
        Set wb = Workbooks.Open(SETTINGS.Range("InputTemplate").text, False, False)
        If Err.Number <> 0 Then
            Err.Number = 0
            messageBox "Template File Does not Exist" & vbNewLine & vbNewLine & "Select a valid `Calc. Template Excel` file in the `Settings` section of `Bulk-Calculate` menu", "Missing Template File", vbCritical
            Unload Me
        End If
    End If
    Me.SheetNamePicker.Clear
    For Each ws In wb.Sheets
        If ws.visible = xlSheetVisible Then
            Me.SheetNamePicker.AddItem ws.Name
        End If
    Next ws
    ThisWorkbook.Activate
    Me.CellNameCheckBox.Value = SETTINGS.Range("UseCellNameInsteadOfAddress").Value
    Me.AutomaticallyTriggerCellPicker.Value = SETTINGS.Range("AutomaticallyTriggerCellPicker").Value
    Me.MultiPage.Value = 0
    If Me.AutomaticallyTriggerCellPicker Then
        MappedCellPickerButton_Click
        If Me.CellReference <> "" Then
            MappedVariablePickerButton_Click
        End If
    End If
End Sub

