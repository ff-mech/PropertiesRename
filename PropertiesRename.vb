Option Explicit

' ============================================================================
'  PropertiesRename.vb
'  SOLIDWORKS Macro  -  Batch Property & Revision Table Update
' ----------------------------------------------------------------------------
'  Scans a folder for SOLIDWORKS files and updates the following:
'    Parts / Assemblies : Sets the "DrawnBy" custom property.
'    Drawings           : Enforces all 15 standard custom properties,
'                         sets "DwgDrawnBy", and inserts Rev A with
'                         the description "INITIAL RELEASE".
'
'  Entry point : main()
' ============================================================================

Dim swApp As SldWorks.SldWorks

Private Const REVISION_TEMPLATE As String = "\\npsvr05\FOXFAB\FOXFAB_DATA\ENGINEERING\SOLIDWORKS\Foxfab Templates\Revision Table v1.1.sldrevtbt"


' ============================================================================
'  ShowInputDialog
'  Displays an HTA dialog to collect all required user inputs.
'  Returns True if the user confirmed, False if cancelled.
' ============================================================================
Function ShowInputDialog(ByRef targetFolder    As String, _
                         ByRef modelInitials   As String, _
                         ByRef drawingInitials As String, _
                         ByRef skip003         As Boolean) As Boolean

    ShowInputDialog = False

    Dim htaPath    As String
    Dim resultPath As String
    htaPath    = Environ("TEMP") & "\SW_BatchUpdate.hta"
    resultPath = Environ("TEMP") ;& "\SW_BatchUpdate_result.txt"

    ' Remove any leftover result file from a previous run
    On Error Resume Next
    Kill resultPath
    On Error GoTo 0

    ' ----- Build the HTA GUI -----
    Dim f As Integer
    f = FreeFile
    Open htaPath For Output As #f

    Print #f, "<html>"
    Print #f, "<head>"
    Print #f, "<title>DrawnBy Batch Update</title>"
    Print #f, "<HTA:APPLICATION"
    Print #f, "  ID=""BatchUpdate"""
    Print #f, "  APPLICATIONNAME=""DrawnBy Batch Update"""
    Print #f, "  BORDER=""dialog"""
    Print #f, "  BORDERSTYLE=""normal"""
    Print #f, "  INNERBORDER=""no"""
    Print #f, "  CAPTION=""yes"""
    Print #f, "  MAXIMIZEBUTTON=""no"""
    Print #f, "  MINIMIZEBUTTON=""no"""
    Print #f, "  SYSMENU=""yes"""
    Print #f, "  SCROLL=""no"""
    Print #f, "  SINGLEINSTANCE=""yes"""
    Print #f, "  WINDOWSTATE=""normal"">"
    Print #f, "<style>"
    Print #f, "  body { font-family: Segoe UI, Tahoma, sans-serif; font-size: 9pt; margin: 15px; background: #f0f0f0; }"
    Print #f, "  h2 { margin: 0 0 12px 0; color: #333; font-size: 13pt; }"
    Print #f, "  label { display: block; margin: 8px 0 3px 0; font-weight: bold; color: #444; }"
    Print #f, "  input[type=text] { width: 100%; padding: 5px; font-size: 9pt; border: 1px solid #999; box-sizing: border-box; }"
    Print #f, "  .chkrow { margin: 12px 0; }"
    Print #f, "  .chkrow input { vertical-align: middle; }"
    Print #f, "  .chkrow label { display: inline; font-weight: normal; margin-left: 4px; }"
    Print #f, "  .buttons { text-align: right; margin-top: 15px; padding-top: 10px; border-top: 1px solid #ccc; }"
    Print #f, "  .buttons button { padding: 6px 20px; font-size: 9pt; margin-left: 8px; cursor: pointer; }"
    Print #f, "  .btnOK { background: #0078d4; color: white; border: 1px solid #005a9e; }"
    Print #f, "  .btnOK:hover { background: #005a9e; }"
    Print #f, "  .btnCancel { background: #e0e0e0; border: 1px solid #999; }"
    Print #f, "</style>"
    Print #f, "<script language=""VBScript"">"
    Print #f, "  Sub Window_OnLoad"
    Print #f, "    window.resizeTo 440, 370"
    Print #f, "    Dim sl, st"
    Print #f, "    sl = (screen.width  - 440) / 2"
    Print #f, "    st = (screen.height - 370) / 2"
    Print #f, "    window.moveTo sl, st"
    Print #f, "    document.getElementById(""txtFolder"").focus"
    Print #f, "  End Sub"
    Print #f, ""
    Print #f, "  Sub btnOK_Click"
    Print #f, "    Dim fso, f"
    Print #f, "    Set fso = CreateObject(""Scripting.FileSystemObject"")"
    Print #f, "    Set f = fso.CreateTextFile(""" & Replace(resultPath, "\", "\\") & """, True)"
    Print #f, "    f.WriteLine document.getElementById(""txtFolder"").value"
    Print #f, "    f.WriteLine document.getElementById(""txtModel"").value"
    Print #f, "    f.WriteLine document.getElementById(""txtDraw"").value"
    Print #f, "    If document.getElementById(""chkSkip"").checked Then"
    Print #f, "      f.WriteLine ""YES"""
    Print #f, "    Else"
    Print #f, "      f.WriteLine ""NO"""
    Print #f, "    End If"
    Print #f, "    f.Close"
    Print #f, "    self.close"
    Print #f, "  End Sub"
    Print #f, ""
    Print #f, "  Sub btnCancel_Click"
    Print #f, "    self.close"
    Print #f, "  End Sub"
    Print #f, ""
    Print #f, "  Sub CheckEnter()"
    Print #f, "    If window.event.keyCode = 13 Then btnOK_Click"
    Print #f, "  End Sub"
    Print #f, "</script>"
    Print #f, "</head>"
    Print #f, "<body onkeypress=""CheckEnter"">"
    Print #f, "<h2>Batch DrawnBy / Property Update</h2>"
    Print #f, ""
    Print #f, "<label for=""txtFolder"">Folder Path:</label>"
    Print #f, "<input type=""text"" id=""txtFolder"" value="""">"
    Print #f, ""
    Print #f, "<label for=""txtModel"">DrawnBy Initials (parts / assemblies):</label>"
    Print #f, "<input type=""text"" id=""txtModel"" value="""">"
    Print #f, ""
    Print #f, "<label for=""txtDraw"">DwgDrawnBy Initials (drawings):</label>"
    Print #f, "<input type=""text"" id=""txtDraw"" value="""">"
    Print #f, ""
    Print #f, "<div class=""chkrow"">"
    Print #f, "  <input type=""checkbox"" id=""chkSkip"">"
    Print #f, "  <label for=""chkSkip"">Skip files starting with '003-'</label>"
    Print #f, "</div>"
    Print #f, ""
    Print #f, "<div class=""buttons"">"
    Print #f, "  <button class=""btnCancel"" onclick=""btnCancel_Click"">Cancel</button>"
    Print #f, "  <button class=""btnOK""     onclick=""btnOK_Click"">Run</button>"
    Print #f, "</div>"
    Print #f, ""
    Print #f, "</body>"
    Print #f, "</html>"

    Close #f

    ' Launch the HTA and block until it closes
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & htaPath & """", 1, True

    ' If no result file was created, the user cancelled
    If Dir(resultPath) = "" Then
        ShowInputDialog = False
        Exit Function
    End If

    ' ----- Read the results written by the HTA -----
    Dim resultLines(1 To 4) As String
    Dim lineNum As Long
    lineNum = 0

    f = FreeFile
    Open resultPath For Input As #f
    Do While Not EOF(f) And lineNum < 4
        lineNum = lineNum + 1
        Line Input #f, resultLines(lineNum)
    Loop
    Close #f

    ' Clean up temp files
    On Error Resume Next
    Kill htaPath
    Kill resultPath
    On Error GoTo 0

    If lineNum < 4 Then
        ShowInputDialog = False
        Exit Function
    End If

    targetFolder    = Trim$(resultLines(1))
    modelInitials   = Trim$(resultLines(2))
    drawingInitials = Trim$(resultLines(3))
    skip003         = (UCase$(Trim$(resultLines(4))) = "YES")

    ' All three text fields are required
    If targetFolder = "" Or modelInitials = "" Or drawingInitials = "" Then
        ShowInputDialog = False
        Exit Function
    End If

    ShowInputDialog = True

End Function


' ============================================================================
'  main
'  Entry point. Runs the full batch update in four phases:
'    1. Collect user input
'    2. Scan the target folder
'    3. Process each file
'    4. Write the log and report results
' ============================================================================
Sub main()

    Set swApp = Application.SldWorks

    ' --------------------------------------------------------------------------
    '  Phase 1 - Collect User Input
    ' --------------------------------------------------------------------------

    Dim targetFolder    As String
    Dim modelInitials   As String
    Dim drawingInitials As String
    Dim skip003         As Boolean

    If Not ShowInputDialog(targetFolder, modelInitials, drawingInitials, skip003) Then
        Exit Sub
    End If

    ' Ensure the folder path ends with a backslash
    If Right$(targetFolder, 1) <> "\" Then targetFolder = targetFolder & "\"

    Dim folderOK As Boolean
    folderOK = False
    On Error Resume Next
    folderOK = ((GetAttr(targetFolder) And vbDirectory) = vbDirectory)
    On Error GoTo 0

    If Not folderOK Then
        swApp.SendMsgToUser2 "Folder does not exist:" & vbCrLf & targetFolder, swMbStop, swMbOk
        Exit Sub
    End If

    ' --------------------------------------------------------------------------
    '  Phase 2 - Scan Files
    ' --------------------------------------------------------------------------

    Dim jPaths()   As String
    Dim jNames()   As String
    Dim jTypes()   As Long
    Dim jIsDrw()   As Boolean
    Dim jobCount   As Long
    Dim partCount  As Long
    Dim asmCount   As Long
    Dim drwCount   As Long
    Dim skipped003 As Long

    jobCount = 0 : partCount = 0 : asmCount = 0 : drwCount = 0 : skipped003 = 0

    Dim fName As String
    fName = Dir(targetFolder & "*.*")

    Do While fName <> ""
        Dim uName As String
        uName = UCase$(fName)

        If skip003 And Left$(fName, 4) = "003-" Then
            skipped003 = skipped003 + 1
            GoTo SkipFile
        End If

        Dim dt    As Long
        Dim isDrw As Boolean
        dt = -1

        Select Case True
            Case Right$(uName, 7) = ".SLDPRT" : dt = swDocPART     : isDrw = False : partCount = partCount + 1
            Case Right$(uName, 7) = ".SLDASM" : dt = swDocASSEMBLY : isDrw = False : asmCount  = asmCount  + 1
            Case Right$(uName, 7) = ".SLDDRW" : dt = swDocDRAWING  : isDrw = True  : drwCount  = drwCount  + 1
        End Select

        If dt >= 0 Then
            jobCount = jobCount + 1
            ReDim Preserve jPaths(1 To jobCount)
            ReDim Preserve jNames(1 To jobCount)
            ReDim Preserve jTypes(1 To jobCount)
            ReDim Preserve jIsDrw(1 To jobCount)
            jPaths(jobCount) = targetFolder ;& fName
            jNames(jobCount) = fName
            jTypes(jobCount) = dt
            jIsDrw(jobCount) = isDrw
        End If

SkipFile:
        fName = Dir
    Loop

    If jobCount = 0 Then
        swApp.SendMsgToUser2 _
            "No SOLIDWORKS files found in:" & vbCrLf & targetFolder & _
            IIf(skipped003 > 0, vbCrLf & vbCrLf & "(Skipped " & skipped003 & " file(s) starting with '003-')", ""), _
            swMbWarning, swMbOk
        Exit Sub
    End If

    Dim confirmMsg As String
    confirmMsg = "Folder: " & targetFolder & vbCrLf & vbCrLf & _
                 "Files to process:"                          & vbCrLf & _
                 "  Parts:      " & partCount                  & vbCrLf & _
                 "  Assemblies: " & asmCount                   & vbCrLf & _
                 "  Drawings:   " & drwCount                   & vbCrLf & vbCrLf & _
                 "DrawnBy    (parts / asm) = " & modelInitials    & vbCrLf & _
                 "DwgDrawnBy (drawings)    = " & drawingInitials  & vbCrLf & _
                 "Skip 003-  files         = " & IIf(skip003, "YES (" & skipped003 & " skipped)", "NO") & vbCrLf & vbCrLf & _
                 "Proceed?"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Batch Update") <> vbYes Then Exit Sub

    ' --------------------------------------------------------------------------
    '  Phase 3 - Process Each File
    ' --------------------------------------------------------------------------

    Dim logPath      As String
    logPath = targetFolder & "DrawnBy_Update_Log.txt"

    Dim overallStart As Single
    overallStart = Timer

    Dim updatedCount As Long
    Dim skippedCount As Long
    Dim failedCount  As Long
    Dim updateLog    As String
    Dim skipLog      As String
    Dim failLog      As String
    Dim debugLog     As String

    updatedCount = 0 : skippedCount = 0 : failedCount = 0
    updateLog = "" : skipLog = "" : failLog = "" : debugLog = ""

    Dim i As Long
    For i = 1 To jobCount

        Dim filePath  As String
        Dim fileName  As String
        Dim docType   As Long
        Dim isDrawing As Boolean

        filePath  = jPaths(i)
        fileName  = jNames(i)
        docType   = jTypes(i)
        isDrawing = jIsDrw(i)

        debugLog = debugLog & "--- " & fileName & " (type=" & docType & ", isDrw=" & isDrawing & ") ---" & vbCrLf

        ' Skip read-only files
        Dim isRO As Boolean
        isRO = False
        On Error Resume Next
        isRO = ((GetAttr(filePath) And vbReadOnly) = vbReadOnly)
        On Error GoTo 0

        If isRO Then
            skipLog      = skipLog      & "SKIP - Read-only: " & fileName & vbCrLf
            skippedCount = skippedCount + 1
            debugLog     = debugLog     & "  -> Skipped (read-only)" & vbCrLf
            GoTo NextJob
        End If

        ' Close the file in SOLIDWORKS if it is already open
        On Error Resume Next
        Dim vDocs As Variant
        vDocs = swApp.GetDocuments
        If Not IsEmpty(vDocs) Then
            Dim d As Long
            For d = 0 To UBound(vDocs)
                Dim tmpDoc As SldWorks.ModelDoc2
                Set tmpDoc = vDocs(d)
                If Not tmpDoc Is Nothing Then
                    If LCase$(tmpDoc.GetPathName) = LCase$(filePath) Then
                        swApp.CloseDoc tmpDoc.GetTitle
                        Exit For
                    End If
                End If
            Next d
        End If
        On Error GoTo 0

        ' Open the file silently
        Dim openErrors   As Long
        Dim openWarnings As Long
        Dim swModel      As SldWorks.ModelDoc2
        Set swModel = swApp.OpenDoc6(filePath, docType, swOpenDocOptions_Silent, "", openErrors, openWarnings)

        If swModel Is Nothing Then
            failLog     = failLog     & "FAIL - Could not open: " & fileName & " (error code: " & openErrors & ")" & vbCrLf
            failedCount = failedCount + 1
            debugLog    = debugLog    & "  -> FAIL: OpenDoc6 returned Nothing (error=" & openErrors & ", warnings=" & openWarnings & ")" & vbCrLf
            GoTo NextJob
        End If

        debugLog = debugLog & "  Opened: " & swModel.GetTitle & vbCrLf

        Dim partLog   As String
        Dim didChange As Boolean
        partLog   = ""
        didChange = False

        ' ---- Drawing: enforce all 15 properties + add revision table ----
        If isDrawing Then

            Dim enfResult As Boolean
            enfResult = EnforceDrawingProperties(swModel, drawingInitials, partLog, didChange)
            debugLog  = debugLog & "  EnforceDrawingProperties = " & enfResult & vbCrLf
            If partLog <> "" Then debugLog = debugLog & partLog

            If Not enfResult Then
                failLog     = failLog     ;& "FAIL - Property enforcement failed: " ;& fileName ;& vbCrLf ;& partLog
                failedCount = failedCount + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If

            Dim revResult As Boolean
            revResult = UpdateDrawingRevisionTable(swModel, fileName, partLog, didChange)
            debugLog  = debugLog ;& "  UpdateDrawingRevisionTable = " ;& revResult ;& vbCrLf
            If partLog <> "" Then debugLog = debugLog ;& partLog

            If Not revResult Then
                failLog     = failLog     ;& "FAIL - Revision table update failed: " ;& fileName ;& vbCrLf ;& partLog
                failedCount = failedCount + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If

        ' ---- Part / Assembly: set DrawnBy property only ----
        Else

            Dim propResult As Boolean
            propResult = SetOrCreateCustomPropertyIfNeeded(swModel, "DrawnBy", modelInitials, didChange)
            debugLog   = debugLog & "  SetOrCreateCustomPropertyIfNeeded = " & propResult & ", didChange=" & didChange & vbCrLf

            If Not propResult Then
                failLog     = failLog     & "FAIL - Could not set DrawnBy: " & fileName & vbCrLf
                failedCount = failedCount + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If

        End If

        ' ---- Save if any changes were made ----
        If didChange Then

            Dim saveErrors   As Long
            Dim saveWarnings As Long
            swModel.SetSaveFlag
            Call swModel.Save3(swSaveAsOptions_Silent, saveErrors, saveWarnings)

            debugLog = debugLog ;& "  Save3: errors=" ;& saveErrors ;& ", warnings=" ;& saveWarnings ;& vbCrLf

            If saveErrors <> 0 Then
                failLog     = failLog     ;& "FAIL - Could not save: " ;& fileName ;& " (save error: " ;& saveErrors ;& ")" ;& vbCrLf
                failedCount = failedCount + 1
                swApp.CloseDoc swModel.GetTitle
                GoTo NextJob
            End If

            updateLog = updateLog ;& "UPDATED - " ;& fileName
            If isDrawing Then
                updateLog = updateLog ;& " | 15 properties enforced | DwgDrawnBy=" ;& drawingInitials ;& _
                                        " | Revision=A | Description=INITIAL RELEASE"
            Else
                updateLog = updateLog ;& " | DrawnBy=" ;& modelInitials
            End If
            updateLog    = updateLog    ;& vbCrLf
            updatedCount = updatedCount + 1
            debugLog     = debugLog     ;& "  -> UPDATED" ;& vbCrLf

        Else
            skipLog      = skipLog      ;& "SKIP - No change needed: " ;& fileName ;& vbCrLf
            skippedCount = skippedCount + 1
            debugLog     = debugLog     ;& "  -> Skipped (no change)" ;& vbCrLf
        End If

        swApp.CloseDoc swModel.GetTitle

NextJob:
    Next i

    ' --------------------------------------------------------------------------
    '  Phase 4 - Write Log
    ' --------------------------------------------------------------------------

    Dim elapsed As Double
    elapsed = Timer - overallStart
    If elapsed < 0 Then elapsed = elapsed + 86400   ' Handle midnight rollover

    Dim hh As Long
    Dim mm As Long
    Dim ss As Long
    hh = Int(elapsed / 3600)
    mm = Int((elapsed - hh * 3600) / 60)
    ss = Int(elapsed - hh * 3600 - mm * 60)

    Dim elapsedStr As String
    elapsedStr = Format$(hh, "00") & ":" & Format$(mm, "00") & ":" & Format$(ss, "00")

    Dim lf As Integer
    lf = FreeFile
    Open logPath For Output As #lf

    Print #lf, "DRAWNBY / REVISION BATCH UPDATE LOG"
    Print #lf, "Started    : " & Now
    Print #lf, "Folder     : " & targetFolder
    Print #lf, "DrawnBy    : " & modelInitials
    Print #lf, "DwgDrawnBy : " & drawingInitials
    Print #lf, "Skip 003-  : " & IIf(skip003, "YES (" & skipped003 & " skipped)", "NO")
    Print #lf, "Template   : " & REVISION_TEMPLATE
    Print #lf, String$(70, "=")
    Print #lf, ""

    Print #lf, "FAILURES"
    Print #lf, String$(70, "-")
    If failLog  <> "" Then Print #lf, failLog  Else Print #lf, "None" & vbCrLf

    Print #lf, "SKIPPED"
    Print #lf, String$(70, "-")
    If skipLog  <> "" Then Print #lf, skipLog  Else Print #lf, "None" & vbCrLf

    Print #lf, "UPDATED"
    Print #lf, String$(70, "-")
    If updateLog <> "" Then Print #lf, updateLog Else Print #lf, "None" & vbCrLf

    Print #lf, "SUMMARY"
    Print #lf, String$(70, "-")
    Print #lf, "  Updated  : " & updatedCount
    Print #lf, "  Skipped  : " & skippedCount
    Print #lf, "  Failed   : " & failedCount
    Print #lf, "  Total    : " & jobCount
    Print #lf, "  Elapsed  : " & elapsedStr
    Print #lf, String$(70, "=")
    Print #lf, ""

    Print #lf, "DEBUG LOG"
    Print #lf, String$(70, "=")
    Print #lf, debugLog

    Close #lf

    swApp.SendMsgToUser2 _
        "Done!" & vbCrLf & vbCrLf & _
        "Updated : " & updatedCount & vbCrLf & _
        "Skipped : " & skippedCount & vbCrLf & _
        "Failed  : " & failedCount  & vbCrLf & _
        "Elapsed : " & elapsedStr   & vbCrLf & vbCrLf & _
        "Log saved to:" & vbCrLf & logPath, _
        swMbInformation, swMbOk

End Sub


' ============================================================================
'  SetOrCreateCustomPropertyIfNeeded
'  Sets propName to propValue on the model's custom property manager.
'  Skips the write if the property already holds the correct value.
'  Sets didChange = True only when a write is performed.
'  Returns True on success, False on error.
' ============================================================================
Function SetOrCreateCustomPropertyIfNeeded(ByVal swModel   As SldWorks.ModelDoc2, _
                                           ByVal propName  As String, _
                                           ByVal propValue As String, _
                                           ByRef didChange As Boolean) As Boolean

    On Error GoTo EH

    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")

    Dim valOut         As String
    Dim resolvedValOut As String
    Dim wasResolved    As Boolean
    Dim linkToProp     As Boolean

    valOut = "" : resolvedValOut = ""

    Call swCustPropMgr.Get6(propName, False, valOut, resolvedValOut, wasResolved, linkToProp)

    ' Property already has the correct value - nothing to do
    If StrComp(Trim$(resolvedValOut), Trim$(propValue), vbTextCompare) = 0 Or _
       StrComp(Trim$(valOut),         Trim$(propValue), vbTextCompare) = 0 Then
        SetOrCreateCustomPropertyIfNeeded = True
        Exit Function
    End If

    ' Try Add3 first; fall back to Set2 if the property already exists
    Dim addResult As Long
    addResult = swCustPropMgr.Add3(propName, swCustomInfoText, propValue, swCustomPropertyDeleteAndAdd)

    If addResult >= 0 Then
        didChange = True
        SetOrCreateCustomPropertyIfNeeded = True
    Else
        SetOrCreateCustomPropertyIfNeeded = (swCustPropMgr.Set2(propName, propValue) <> 0)
        If SetOrCreateCustomPropertyIfNeeded Then didChange = True
    End If

    Exit Function

EH:
    SetOrCreateCustomPropertyIfNeeded = False

End Function


' ============================================================================
'  EnforceDrawingProperties
'  Ensures the drawing holds exactly the 15 standard custom properties in the
'  correct order, preserving existing values and updating DwgDrawnBy.
'  Sets didChange = True when any modification is made.
'  Returns True on success, False on error.
' ============================================================================
Function EnforceDrawingProperties(ByVal swModel         As SldWorks.ModelDoc2, _
                                  ByVal dwgDrawnByValue As String, _
                                  ByRef partLog         As String, _
                                  ByRef didChange       As Boolean) As Boolean

    On Error GoTo EH

    Dim swCustPropMgr As SldWorks.CustomPropertyManager
    Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")

    If swCustPropMgr Is Nothing Then
        partLog = partLog ;& "  Could not get CustomPropertyManager." ;& vbCrLf
        EnforceDrawingProperties = False
        Exit Function
    End If

    ' Snapshot all 15 current property values before making changes
    Dim existingValues(1 To 15)      As String
    Dim existingExpressions(1 To 15) As String
    Dim propIdx As Long

    For propIdx = 1 To 15
        Dim pName          As String
        Dim valOut         As String
        Dim resolvedValOut As String
        Dim wasResolved    As Boolean
        Dim linkToProp     As Boolean

        pName = GetDrawingPropertyName(propIdx)
        valOut = "" : resolvedValOut = ""

        Call swCustPropMgr.Get6(pName, False, valOut, resolvedValOut, wasResolved, linkToProp)

        existingExpressions(propIdx) = valOut
        existingValues(propIdx)      = resolvedValOut
    Next propIdx

    ' Check whether the property list is already in the correct state
    Dim vNames As Variant
    vNames = swCustPropMgr.GetNames

    Dim alreadyCorrect As Boolean
    alreadyCorrect = True

    If IsEmpty(vNames) Then
        alreadyCorrect = False
    ElseIf UBound(vNames) - LBound(vNames) + 1 <> 15 Then
        alreadyCorrect = False
    Else
        Dim chkIdx As Long
        For chkIdx = 0 To UBound(vNames)
            If StrComp(CStr(vNames(chkIdx)), GetDrawingPropertyName(chkIdx + 1), vbTextCompare) <> 0 Then
                alreadyCorrect = False
                Exit For
            End If
        Next chkIdx
    End If

    Dim dwgDrawnByCurrent As String
    dwgDrawnByCurrent = Trim$(existingValues(7))
    If dwgDrawnByCurrent = "" Then dwgDrawnByCurrent = Trim$(existingExpressions(7))

    Dim dwgDrawnByMatch As Boolean
    dwgDrawnByMatch = (StrComp(dwgDrawnByCurrent, Trim$(dwgDrawnByValue), vbTextCompare) = 0)

    ' Nothing to do if already correct
    If alreadyCorrect And dwgDrawnByMatch Then
        EnforceDrawingProperties = True
        Exit Function
    End If

    ' Remove all existing properties so they can be re-added in the correct order
    If Not IsEmpty(vNames) Then
        Dim delIdx As Long
        For delIdx = LBound(vNames) To UBound(vNames)
            swCustPropMgr.Delete2 CStr(vNames(delIdx))
        Next delIdx
    End If

    ' Re-add all 15 properties in order, updating DwgDrawnBy (index 7)
    For propIdx = 1 To 15
        pName = GetDrawingPropertyName(propIdx)

        Dim valueToSet As String
        If propIdx = 7 Then
            valueToSet = dwgDrawnByValue           ' Always write the new initials
        Else
            valueToSet = existingExpressions(propIdx)  ' Preserve existing expression
        End If

        Dim addResult As Long
        addResult = swCustPropMgr.Add3(pName, swCustomInfoText, valueToSet, swCustomPropertyOnlyIfNew)

        If addResult < 0 Then
            partLog = partLog & "  WARN - Could not add property: " & pName & " (result: " & addResult & ")" & vbCrLf
        End If
    Next propIdx

    didChange = True
    EnforceDrawingProperties = True
    Exit Function

EH:
    partLog = partLog & "  Exception in EnforceDrawingProperties: " & Err.Description & vbCrLf
    EnforceDrawingProperties = False

End Function


' ============================================================================
'  GetDrawingPropertyName
'  Returns the standard custom property name for a given 1-based index.
'  The 15 properties are defined in the order required by the drawing template.
' ============================================================================
Function GetDrawingPropertyName(ByVal idx As Long) As String
    Select Case idx
        Case 1  : GetDrawingPropertyName = "SWFormatSize"
        Case 2  : GetDrawingPropertyName = "Revision"
        Case 3  : GetDrawingPropertyName = "Description"
        Case 4  : GetDrawingPropertyName = "Material"
        Case 5  : GetDrawingPropertyName = "Finish"
        Case 6  : GetDrawingPropertyName = "DrawnBy"
        Case 7  : GetDrawingPropertyName = "DwgDrawnBy"
        Case 8  : GetDrawingPropertyName = "Bend Deduction"
        Case 9  : GetDrawingPropertyName = "Top Die"
        Case 10 : GetDrawingPropertyName = "Bottom Die"
        Case 11 : GetDrawingPropertyName = "Tol X"
        Case 12 : GetDrawingPropertyName = "Tol X.X"
        Case 13 : GetDrawingPropertyName = "Tol X.XX"
        Case 14 : GetDrawingPropertyName = "Tol X.XXX"
        Case 15 : GetDrawingPropertyName = Chr$(84) ;& Chr$(111) ;& Chr$(108) ;& Chr$(32) ;& Chr$(176)  ' "Tol ?"
        Case Else : GetDrawingPropertyName = ""
    End Select
End Function


' ============================================================================
'  UpdateDrawingRevisionTable
'  Inserts or reuses a revision table on the first sheet of the drawing,
'  clears any existing revision rows, then adds Rev A / INITIAL RELEASE.
'  Sets didChange = True when changes are applied.
'  Returns True on success, False on error.
' ============================================================================
Function UpdateDrawingRevisionTable(ByVal swModel   As SldWorks.ModelDoc2, _
                                    ByVal fileName  As String, _
                                    ByRef partLog   As String, _
                                    ByRef didChange As Boolean) As Boolean

    On Error GoTo EH

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel

    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames

    If IsEmpty(vSheetNames) Then
        partLog = partLog & "  No sheets found." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    ' Activate the first sheet
    Dim firstSheetName As String
    firstSheetName = CStr(vSheetNames(LBound(vSheetNames)))

    Dim swSheet As SldWorks.Sheet

    If swDraw.ActivateSheet(firstSheetName) Then
        Set swSheet = swDraw.GetCurrentSheet
    Else
        Set swSheet = swDraw.GetCurrentSheet
        If swSheet Is Nothing Then
            partLog = partLog & "  Could not activate first sheet: " & firstSheetName & vbCrLf
            UpdateDrawingRevisionTable = False
            Exit Function
        End If
        partLog = partLog & "  WARN - Could not activate '" & firstSheetName & "'; using current: " & swSheet.GetName & vbCrLf
    End If

    If swSheet Is Nothing Then
        partLog = partLog & "  Could not get current sheet." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    ' Get or insert revision table
    Dim swRevTable       As SldWorks.RevisionTableAnnotation
    Dim insertedNewTable As Boolean
    insertedNewTable = False

    Set swRevTable = swSheet.RevisionTable

    If swRevTable Is Nothing Then
        Dim tExists As Boolean
        tExists = False
        On Error Resume Next
        tExists = ((GetAttr(REVISION_TEMPLATE) And vbDirectory) = 0)
        If Err.Number <> 0 Then tExists = False
        On Error GoTo EH

        If Not tExists Then
            partLog = partLog & "  Revision template not found: " & REVISION_TEMPLATE & vbCrLf
            UpdateDrawingRevisionTable = False
            Exit Function
        End If

        Set swRevTable = swSheet.InsertRevisionTable(True, 0#, 0#, swBOMConfigurationAnchor_TopRight, REVISION_TEMPLATE)

        If swRevTable Is Nothing Then
            partLog = partLog & "  InsertRevisionTable failed." & vbCrLf
            UpdateDrawingRevisionTable = False
            Exit Function
        End If

        insertedNewTable = True
        partLog = partLog & "  INFO - Revision table inserted from template." & vbCrLf
    End If

    Dim swTable As SldWorks.TableAnnotation
    Set swTable = swRevTable

    If swTable Is Nothing Then
        partLog = partLog & "  Revision table annotation unavailable." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    ' Scan all rows to find the DESCRIPTION column (accounts for a title row)
    Dim descCol As Long
    Dim r       As Long
    Dim c       As Long
    descCol = -1

    For r = 0 To swTable.RowCount - 1
        For c = 0 To swTable.ColumnCount - 1
            If UCase$(Trim$(swTable.Text2(r, c, False))) = "DESCRIPTION" Then
                descCol = c
                Exit For
            End If
        Next c
        If descCol >= 0 Then Exit For
    Next r

    If descCol < 0 Then
        partLog = partLog & "  Could not find DESCRIPTION column. Table dump:" & vbCrLf
        For r = 0 To swTable.RowCount - 1
            For c = 0 To swTable.ColumnCount - 1
                partLog = partLog & "    [" & r & "," & c & "] = """ & swTable.Text2(r, c, False) & """" & vbCrLf
            Next c
        Next r
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    partLog = partLog & "  DESCRIPTION found at column " & descCol & " (header row " & r & ")" & vbCrLf

    ' Clear existing revision rows if the table was already present
    If Not insertedNewTable Then
        Dim rowIdx As Long
        Dim revId  As Long
        For rowIdx = swTable.RowCount - 1 To 2 Step -1
            revId = swRevTable.GetIdForRowNumber(rowIdx)
            If revId <> -1 Then swRevTable.DeleteRevision revId, True
        Next rowIdx
    End If

    ' Add Rev A
    Dim newRevId As Long
    newRevId = swRevTable.AddRevision("A")

    If newRevId < 0 Then
        partLog = partLog & "  Could not add revision A." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    Dim newRow As Long
    newRow = swRevTable.GetRowNumberForId(newRevId)

    If newRow < 0 Then
        partLog = partLog & "  Could not locate new revision row (id=" & newRevId & ")." & vbCrLf
        UpdateDrawingRevisionTable = False
        Exit Function
    End If

    swTable.Text2(newRow, descCol, True) = "INITIAL RELEASE"
    partLog = partLog & "  Set DESCRIPTION='INITIAL RELEASE' at [" & newRow & "," & descCol & "]" & vbCrLf

    didChange = True
    UpdateDrawingRevisionTable = True
    Exit Function

EH:
    partLog = partLog & "  Exception in UpdateDrawingRevisionTable: " & Err.Description & vbCrLf
    UpdateDrawingRevisionTable = False

End Function
