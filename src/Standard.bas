Attribute VB_Name = "Standard"
' Include reference to `Microsoft Scripting Runtime` by visiting Tools => References

' <<<<<<<<<<<<<<<< Convienience Functions >>>>>>>>>>>>>>>>

Function AppName() As String
    ' Return the application name.
    AppName = "Bulk-Calculate"
End Function

Public Function item_lookup(ByVal target As Variant, lkuprng As Range, ByVal lkupcol As Integer, Optional ByVal errval As Variant = "True") As Variant
    On Error Resume Next
    ' Default return value in case of an error.
    If errval = "True" Then
        item_lookup = target
    Else
        item_lookup = errval
    End If
    
    ' Perform VLOOKUP and return the result.
    item_lookup = Application.WorksheetFunction.VLookup(target, lkuprng, lkupcol, False)
End Function

Public Function item_match(ByVal target As Variant, lkuprng As Range, Optional ByVal method As Integer = 0)
    On Error Resume Next
    ' Default return value indicating no match.
    item_match = -1
    
    ' Perform MATCH and return the result.
    item_match = Application.WorksheetFunction.Match(target, lkuprng, method)
End Function

Sub applySort(srtrng As ListObject, Optional ByVal srtHead As Variant = xlYes, Optional ByVal srtCase As Boolean = False, Optional ByVal srtOrient As Variant = xlTopToBottom, Optional ByVal srtMethod As Variant = xlPinYin)
    ' Set sort properties and apply the sort to the specified range.
    srtrng.Sort.Header = srtHead
    srtrng.Sort.MatchCase = srtCase
    srtrng.Sort.Orientation = srtOrient
    srtrng.Sort.SortMethod = srtMethod
    srtrng.Sort.Apply
    srtrng.Sort.SortFields.Clear
End Sub

Function messageBox(ByVal msg As String, Optional ByVal mtitle As String = "", Optional ByVal msty As VbMsgBoxStyle = vbInformation) As Integer
    ' Display a message box with the specified message, title, and style.
    messageBox = MsgBox(msg, msty, AppName & IIf(Len(mtitle) > 0, " > " & mtitle, ""))
End Function

Sub applyLower(rng As Range)
    Dim cel As Range
    ' Convert each cell value in the range to lowercase.
    For Each cel In rng.Cells
        cel.Value = LCase(cel.Value)
    Next
End Sub

' <<<<<<<<<<<<<<<< File and Folder Browser Functions >>>>>>>>>>>>>>>>

Function FileSelect(ByVal title As String, Optional ByVal initial As String = "", _
    Optional ByVal filter As String = "False", Optional ByVal extn As String = "*.*") As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Set the initial path to the user's profile if not specified.
    If initial <> "" Then
        initial = oFSO.GetParentFolderName(initial)
    Else
        initial = Environ("UserProfile")
    End If
    
    ' Show file dialog for file selection.
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = initial
        .title = title
        .AllowMultiSelect = False
        If filter <> "False" Then
            .Filters.Add filter, extn, 1
        End If
        .Show
        ' Return the selected file path.
        If Not .SelectedItems.Count = 0 Then
            FileSelect = .SelectedItems(1)
        End If
    End With
End Function

Function FolderSelect(ByVal title As String, Optional ByVal initial As String = "C:\") As String
    ' Set initial path to C:\ if not specified.
    If initial = "" Then initial = "C:\"
    
    ' Show folder dialog for folder selection.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = initial
        .title = title
        .Show
        ' Return the selected folder path.
        If Not .SelectedItems.Count = 0 Then
            FolderSelect = .SelectedItems(1)
        End If
    End With
End Function

Sub openFolder(ByVal FolderName As String, Optional ByVal focus = vbNormalFocus)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Open the specified folder in Explorer if it exists.
    If oFSO.FolderExists(FolderName) Then
        Shell "C:\WINDOWS\explorer.exe """ & FolderName & "", focus
    Else
        ' Show a message box if the folder does not exist.
        messageBox """" & FolderName & """ does NOT EXISTS!", "Invalid Folder", vbCritical
    End If
End Sub

' <<<<<<<<<<<<<<<< Path Manipulation Functions >>>>>>>>>>>>>>>>

Sub BuildFullPath(ByVal fullpath)
    ' Convert the path to UNC format.
    fullpath = MakeUNC(fullpath, True)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Recursively build the full path if it doesn't exist.
    If Not oFSO.FolderExists(fullpath) Then
        BuildFullPath oFSO.GetParentFolderName(fullpath)
        oFSO.CreateFolder fullpath
    End If
End Sub

Function MakeUNC(ByVal path As String, Optional ByVal suffix As Boolean = False) As String
    ' Convert path to UNC format.
    MakeUNC = IIf(Left(path, 4) = "\\?\", path, "\\?\" & path)
    If suffix Then
        ' Ensure the path ends with a backslash if suffix is True.
        MakeUNC = IIf(Right(MakeUNC, 1) = "\", MakeUNC, MakeUNC & "\")
    Else
        ' Remove trailing backslash if suffix is False.
        MakeUNC = IIf(Right(MakeUNC, 1) = "\", Left(MakeUNC, Len(MakeUNC) - 1), MakeUNC)
    End If
End Function

Function UnMakeUNC(ByVal path As String, Optional ByVal suffix As Boolean = False) As String
    ' Convert UNC path back to normal format.
    UnMakeUNC = IIf(Left(path, 4) = "\\?\", Right(path, Len(path) - 4), path)
    If suffix Then
        ' Ensure the path ends with a backslash if suffix is True.
        UnMakeUNC = IIf(Right(UnMakeUNC, 1) = "\", UnMakeUNC, UnMakeUNC & "\")
    Else
        ' Remove trailing backslash if suffix is False.
        UnMakeUNC = IIf(Right(UnMakeUNC, 1) = "\", Left(UnMakeUNC, Len(UnMakeUNC) - 1), UnMakeUNC)
    End If
End Function

Function pathJoin(ByVal base As String, ByVal addon As String) As String
    ' Join two path strings ensuring correct backslash placement.
    base = base & IIf(Right(base, 1) = "\", "", "\")
    If Len(addon) > 0 Then
        addon = IIf(Left(addon, 1) = "\", Mid(addon, 2, Len(addon) - 1), addon)
    End If
    If Len(addon) > 0 Then
        addon = IIf(Right(addon, 1) = "\", addon, addon & "\")
    End If
    pathJoin = base & addon
End Function

Function GetRelativePath(ByVal basepath As String, ByVal fullpath As String) As String
    ' Convert paths to UNC format.
    basepath = MakeUNC(basepath, True)
    fullpath = MakeUNC(fullpath, False)
    
    ' Return the relative path by subtracting the base path length from the full path.
    GetRelativePath = Right(fullpath, Len(fullpath) - Len(basepath))
End Function

' <<<<<<<<<<<<<<<< Table Manipulation Functions >>>>>>>>>>>>>>>>

Function tableExist(tableDict As Scripting.dictionary) As Boolean
    On Error Resume Next
    ' Check if the table exists on the specified sheet by attempting to get the row count.
    tableExist = CBool(tableDict("_Sheet_").Range(tableDict("_Table_")).Rows.Count)
    If Err.Number <> 0 Then
        ' If an error occurs, set tableExist to False.
        tableExist = False
        Err.Number = 0
    End If
End Function

Function buildTable(tableDict As Scripting.dictionary, Optional ByVal tableFormat As String = "TableStyleMedium9")
    Dim rng As Range
    Dim i, r, c As Long
    On Error Resume Next
    Set sht = tableDict("_Sheet_")
    
    ' Clear the table if it already exists.
    If tableExist(tableDict) Then
        ClearTable tableDict
    End If
    
    ' Set row and column starting positions.
    r = tableDict("_Row_")
    c = tableDict("_Column_")
    
    ' Populate headers for the table based on the dictionary keys.
    For i = 5 To tableDict.Count - 1
        tableDict("_Sheet_").Cells(r, c + tableDict.Items(i)).Value = CStr(tableDict.Keys(i))
    Next
    
    ' Define the range for the table.
    Set rng = tableDict("_Sheet_").Range(tableDict("_Sheet_").Cells(r, c + 1), tableDict("_Sheet_").Cells(r, c + tableDict.Count - 5))
    
    ' Create the table and apply the specified style.
    rng.Worksheet.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = tableDict("_Table_")
    rng.Worksheet.ListObjects(tableDict("_Table_")).TableStyle = tableFormat
    Set tableDict("_Range_") = sht.Range(tableDict("_Table_"))
    
    ' Check for errors and set the return value accordingly.
    If Err.Number <> 0 Then
        buildTable = False
        Err.Number = 0
    Else
        buildTable = True
    End If
    
    ' Show a message if the table creation failed.
    If Not buildTable Then messageBox "Creation of Table " & tableDict("_Table_") & " failed!" & Err.Number, "", vbCritical
End Function

Sub ClearTable(tableDict As Scripting.dictionary)
    On Error Resume Next
    ' Clear all data from the specified table.
    tableDict("_Sheet_").Range(tableDict("_Table_") & "[#All]").Clear
    
    ' Activate the sheet to refresh the display.
    tableDict("_Sheet_").Activate
    ActiveSheet.UsedRange
End Sub

Function InsertRecord(tableDict As Scripting.dictionary) As Variant
    ' Check if the last cell in the table's first column is not empty.
    If Len(tableDict("_Sheet_").Range(tableDict("_Table_")).Cells(tableDict("_Sheet_").Range(tableDict("_Table_")).Rows.Count, 1).Value) <> 0 Then
        ' Add a new row to the table.
        tableDict("_Sheet_").Range(tableDict("_Table_")).ListObject.ListRows.Add AlwaysInsert:=True
    End If
    
    ' Set the table range.
    Set tableDict("_Range_") = tableDict("_Sheet_").Range(tableDict("_Table_"))
    
    ' Return the new row count.
    InsertRecord = tableDict("_Sheet_").Range(tableDict("_Table_")).Rows.Count
End Function

Function getTableRange(tableDict As Scripting.dictionary) As Range
    ' Set and return the range of the specified table.
    Set tableDict("_Range_") = tableDict("_Sheet_").Range(tableDict("_Table_"))
    Set getTableRange = tableDict("_Range_")
End Function

' <<<<<<<<<<<<<<<< Other Amazing Functions >>>>>>>>>>>>>>>>

Sub RobustCopy(ByVal src, ByVal tar, ByVal nme)
    ' Convert paths from UNC format to normal format.
    src = UnMakeUNC(src, False)
    tar = UnMakeUNC(tar, False)
    
    ' Use Robocopy to copy files from src to tar.
    oShell.Run "cmd /c ""robocopy """ & src & """ """ & tar & """ """ & nme & """""", 0, True
End Sub

Sub exeCmd(ByVal cmd As String, Optional ByVal visible As Integer = 0, Optional ByVal wait As Boolean = False)
    ' Execute a command using cmd.exe.
    oShell.Run "cmd.exe /C """ & cmd & """ & exit", visible, wait
End Sub

Public Function DelStringParser(ByVal DelStr As String, ByVal DelimStr As String, ByVal occurance As Integer) As String
    Dim stposi, edposi, i As Integer
    stposi = 0
    edposi = 0
    
    ' Find the start and end positions of the specified occurrence of the delimiter.
    For i = 1 To occurance Step 1
        stposi = edposi
        edposi = InStr(stposi + 1, DelStr, DelimStr, vbBinaryCompare)
    Next
    
    ' Extract and return the substring between the start and end positions.
    If edposi > stposi Then
        DelStringParser = Mid(DelStr, stposi + 1, edposi - stposi - 1)
    ElseIf stposi > 0 And edposi = 0 Then
        edposi = Len(DelStr)
        DelStringParser = Mid(DelStr, stposi + 1, edposi - stposi)
    End If
End Function

' <<<<<<<<<<<<<<<< End of File >>>>>>>>>>>>>>>>

