Attribute VB_Name = "Ribbon"
' Reference Links
'
' https://www.spreadsheet1.com/office-excel-ribbon-imagemso-icons-gallery-page-01.html
' https://www.rondebruin.nl/win/s2/win003.htm
' https://github.com/fernandreu/office-ribbonx-editor/releases/tag/v1.6
' https://www.rondebruin.nl/win/s1/cdo.htm

Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef _
        destination As Any, ByRef source As Any, ByVal length As Long)

Private myRibbon As IRibbonUI
Public Const version As Double = 1.2
Public Const readmeLink As String = "https://github.com/engineered-in/Synthesizer?tab=readme-ov-file#readme"
Public Const changelogLink As String = "https://github.com/engineered-in/Synthesizer/blob/main/CHANGELOG.md"
Public Const licenseLink As String = "https://github.com/engineered-in/Synthesizer?tab=MIT-1-ov-file#readme"
Global latestVersion As Double

Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
    Dim objRibbon As Object
    CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing
End Function

Sub ResetRibbon()
    Set myRibbon = GetRibbon(SETTINGS.Range("AZ1").Value)
    myRibbon.InvalidateControl ("initiateUpdate")
    Err.Number = 0
End Sub

Sub InvalidateControl(ByVal ControlID As String)
    On Error Resume Next
    myRibbon.InvalidateControl (ControlID)
    If Err.Number <> 0 Then
        ResetRibbon
        myRibbon.InvalidateControl (ControlID)
        Err.Number = 0
    End If
End Sub

Sub Ribbon_OnLoad(ByVal ribbon As Office.IRibbonUI)
    Set myRibbon = ribbon
    SETTINGS.Range("AZ1").Value = ObjPtr(ribbon)
    SUMMARY.Activate
    myRibbon.ActivateTab ("Synthesizer")
    latestVersion = CDbl(GetLatestTag())
    InvalidateControl "initiateUpdate"
    SETTINGS.visible = xlSheetVeryHidden
End Sub

'Callback for readmeLink onAction
Sub LaunchReadme(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink readmeLink
End Sub

'Callback for changelogLink onAction
Sub LaunchChangelog(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink changelogLink
End Sub

'Callback for licenseLink onAction
Sub LaunchLicense(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink licenseLink
End Sub

'Callback for mapWizard onAction
Sub ShowMapDataWizard(control As IRibbonControl)
    Dim wb As Workbook
    Dim canOpen As Boolean
    canOpen = True
    On Error Resume Next
    If SETTINGS.Range("InputTemplate").text = "" Then
        canOpen = False
        messageBox "Template File is necessary for mapping." & vbNewLine _
            & "Please select a valid Template File", "Template File not Selected", vbCritical
    Else
        Set oFSO = CreateObject("Scripting.FileSystemObject") ' Create FileSystemObject
        Set wb = Workbooks(oFSO.GetFileName(SETTINGS.Range("InputTemplate").text)) ' Try to set workbook if already open
        
        If Err.Number <> 0 Then ' If error occurs (workbook not open), open the workbook
            Err.Number = 0
            Set wb = Workbooks.Open(SETTINGS.Range("InputTemplate").text, False, False)
            If Err.Number <> 0 Then
                Err.Number = 0
                canOpen = False
                messageBox "Unable to Open Template File" & vbNewLine _
                & "Please select a valid Template File", "Template File Error", vbCritical
            End If
        End If
    End If
    If canOpen Then
        MappingWizard.Show
    End If
    ThisWorkbook.Activate
    MAPPER.Activate

End Sub

'Callback for initSummary onAction
Sub InitializeSummary(control As IRibbonControl)
    initSummary
End Sub

'Callback for outputFolder getText
Sub GetOutputFolder(control As IRibbonControl, ByRef returnedVal)
    returnedVal = SETTINGS.Range("OutputFolder").Value
End Sub

'Callback for outputFolder onChange
Sub UpdateOutputFolder(control As IRibbonControl, text As String)
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If oFSO.FolderExists(text) Then
        SETTINGS.Range("OutputFolder").Value = text
    End If
    InvalidateControl "outputFolder"

End Sub

'Callback for browseOutput onAction
Sub browseOutputFolder(control As IRibbonControl)
    Dim folderPath As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    folderPath = FolderSelect("Choose a Folder to Trace...", SETTINGS.Range("OutputFolder").Value)
    If oFSO.FolderExists(folderPath) Then
        SETTINGS.Range("OutputFolder").Value = folderPath
        InvalidateControl "outputFolder"
    End If
End Sub

'Callback for setPrev onAction
Sub SetPrevItem(control As IRibbonControl)
    Dim updated As Integer
    updated = SETTINGS.Range("SelectedItem").Value - 1
    If updated <= SUMMARY.Range("Summ").Rows.Count And updated > 0 Then
        SETTINGS.Range("SelectedItem").Value = updated
        initializer
        genOne
        destroyer
    End If
    SUMMARY.Activate
    InvalidateControl "selectedItem"
End Sub

'Callback for setNext onAction
Sub SetNextItem(control As IRibbonControl)
    Dim updated As Integer
    updated = SETTINGS.Range("SelectedItem").Value + 1
    If updated <= SUMMARY.Range("Summ").Rows.Count And updated > 0 Then
        SETTINGS.Range("SelectedItem").Value = updated
        initializer
        genOne
        destroyer
    End If
    SUMMARY.Activate
    InvalidateControl "selectedItem"
End Sub

'Callback for exportFormat onAction
Sub GetExportFormatIndex(control As IRibbonControl, id As String, index As Integer)
    SETTINGS.Range("ExportIndex").Value = index + 1
End Sub

'Callback for exportFormat getSelectedItemIndex
Sub SetExportFormatIndex(control As IRibbonControl, ByRef returnedVal)
    returnedVal = SETTINGS.Range("ExportIndex").Value - 1
End Sub

'Callback for openOutput onAction
Sub openOutputFolder(control As IRibbonControl)
    openFolder SETTINGS.Range("OutputFolder").Value
End Sub

'Callback for selectedItem getText
Sub GetCurrentItem(control As IRibbonControl, ByRef returnedVal)
    SUMMARY.Activate
    returnedVal = SETTINGS.Range("SelectedItem").Value
End Sub

'Callback for selectedItem onChange
Sub UpdateCurrentItem(control As IRibbonControl, text As String)
    On Error Resume Next
    If item_match(CInt(text), SUMMARY.Range("Summ[0]"), 0) <> -1 Then
        If Err.Number = 0 Then
            SETTINGS.Range("SelectedItem").Value = CInt(text)
            initializer
            genOne
            destroyer
        End If
    End If
    SUMMARY.Activate
    InvalidateControl "selectedItem"
End Sub

'Callback for computeAll onAction
Sub ComputeAll(control As IRibbonControl)
    Dim i As Integer
    For i = 1 To SUMMARY.Range("Summ").Rows.Count Step 1
        SETTINGS.Range("SelectedItem").Value = i
        SUMMARY.Activate
        InvalidateControl "selectedItem"
        initializer
        genOne
        destroyer
    Next
    SUMMARY.Activate
End Sub

'Callback for exportAll onAction
Sub ExportAll(control As IRibbonControl)
    Dim i As Integer
    For i = 1 To SUMMARY.Range("Summ").Rows.Count Step 1
        SETTINGS.Range("SelectedItem").Value = i
        InvalidateControl "selectedItem"
        initializer
        printOne
        destroyer
    Next
End Sub

'Callback for exportOne onAction
Sub ExportOne(control As IRibbonControl)
    initializer
    printOne
    destroyer
End Sub


'Callback for help onAction
Sub HelpVideo(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink "https://github.com/engineered-in/Synthesizer/wiki"
End Sub

'Callback for feedback onAction
Sub Feedback(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink "mailto:swarup+synthesizer@engineered.co.in?subject=Synthesizer%20-%20Feedback%20-%20reg.&body=Dear%20Swarup,%0D%0A%0D%0APlease%20find%20below%20my%20feedback%20on%20Synthesizer.xlsb%0D%0A%0D%0AFeedback [Positive/Negative]: %0D%0A%0D%0AComments:"
End Sub

'Callback for ImportData onAction
Sub importSpreadsheets(control As IRibbonControl)
    openExcelFiles
End Sub

'Callback for initiateUpdate getVisible
Sub isUpdatable(control As IRibbonControl, ByRef returnedVal)
    returnedVal = latestVersion > version
End Sub

'Callback for initiateUpdate getLabel
Sub getVersionLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Download v" & CStr(latestVersion)
End Sub

'Callback for initiateUpdate onAction
Sub initiateUpdate(control As IRibbonControl)
    messageBox "A download window will be opened on your default browser.", "Download Version " & CStr(latestVersion), vbInformation
    ThisWorkbook.FollowHyperlink "https://github.com/engineered-in/Synthesizer/releases/latest"
End Sub

'Callback for openCalc onAction
Sub openCalculationTemplate(control As IRibbonControl)
    Workbooks.Open SETTINGS.Range("InputTemplate").Value, False
End Sub

'Callback for calculationTemplate getText
Sub GetCalculationTemplate(control As IRibbonControl, ByRef returnedVal)
    returnedVal = SETTINGS.Range("InputTemplate").Value
End Sub

'Callback for calculationTemplate onChange
Sub UpdateCalculationTemplate(control As IRibbonControl, text As String)
    If text = "" Then
        SETTINGS.Range("InputTemplate").Value = text
    Else
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        If oFSO.FileExists(text) Then
            SETTINGS.Range("InputTemplate").Value = text
        End If
    End If
    InvalidateControl "calculationTemplate"
End Sub

'Callback for browseCalculation onAction
Sub browseCalculationTemplate(control As IRibbonControl)
    Dim filePath As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    filePath = FileSelect("Select the Template Excel File...", SETTINGS.Range("InputTemplate").Value, "Excel Files", "*.xls;*.xlsx;*.xlsb;*.xlsm")
    If oFSO.FileExists(filePath) Then
        SETTINGS.Range("InputTemplate").Value = filePath
        InvalidateControl "calculationTemplate"
    End If

End Sub

Function GetLatestTag() As String
    Dim httpRequest As Object
    Dim URL As String
    Dim response As String
    Dim startPos As Integer, endPos As Integer
    
    URL = "https://api.github.com/repos/engineered-in/Synthesizer/releases/latest"
    GetLatestTag = "1.0"
    
    On Error Resume Next
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    With httpRequest
        .Open "GET", URL, False
        .send
    End With
    
    response = httpRequest.responseText

    startPos = InStr(response, """name"":""") + Len("""name"":""")
    If startPos > Len("""name"":""") Then
        endPos = InStr(startPos, response, """", vbTextCompare)
        GetLatestTag = Mid(response, startPos, endPos - startPos)
    Else
        GetLatestTag = "1.0"
    End If
End Function

