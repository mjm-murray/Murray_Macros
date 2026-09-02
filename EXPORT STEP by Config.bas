Option Explicit

' =====================================================================
'  EXPORT STEP by Config
'  ---------------------------------------------------------------
'  Works on 3D files only (.sldprt or .sldasm).
'  Exports the ACTIVE configuration to STEP, named:
'
'      <filename> - <configuration>.STEP
'
'  Output folder is chosen at run time via a folder-browse dialog.
'  Multibody parts can optionally be split into one STEP per body.
' =====================================================================

' --- Constants ---
Const swDocPART          As Long = 1
Const swDocASSEMBLY      As Long = 2
Const swSolidBody        As Long = 0
Const swSaveVerCurrent   As Long = 0
Const swSaveOpSilent     As Long = 1

' Default folder the picker opens to (change to suit your setup)
Const DEFAULT_EXPORT_FOLDER As String = "C:\Nevados Vault\Export"

' Shell BrowseForFolder options
Const BIF_RETURNONLYFSDIRS  As Long = &H1
Const BIF_NEWDIALOGSTYLE    As Long = &H40

' --- Helper: check file exists on disk ---
Private Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

' --- Helper: join a folder and a filename with exactly one backslash ---
Private Function PathJoin(ByVal folder As String, ByVal fileName As String) As String
    If Len(folder) = 0 Then
        PathJoin = fileName
    ElseIf Right$(folder, 1) = "\" Then
        PathJoin = folder & fileName          ' e.g. "C:\" + "file"
    Else
        PathJoin = folder & "\" & fileName
    End If
End Function

' --- Helper: let the user pick an export folder (returns "" if cancelled) ---
Private Function PickExportFolder(ByVal startFolder As String) As String
    Dim shellApp As Object
    Dim folderObj As Object
    Dim chosen   As String

    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    On Error GoTo 0

    If shellApp Is Nothing Then
        ' Shell unavailable for some reason -> fall back to default
        PickExportFolder = startFolder
        Exit Function
    End If

    ' Arg 4 = start folder. If it doesn't exist, Shell just opens at Desktop.
    Set folderObj = shellApp.BrowseForFolder(0&, _
        "Select the folder to export to:", _
        BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE, _
        startFolder)

    If folderObj Is Nothing Then
        PickExportFolder = ""                 ' user cancelled
    Else
        ' .Self.Path gives the filesystem path of the chosen folder
        chosen = folderObj.Self.path
        PickExportFolder = chosen
    End If

    Set folderObj = Nothing
    Set shellApp = Nothing
End Function

' --- Helper: sanitize filename for Windows ---
Private Function SanitizeFileName(ByVal nameIn As String) As String
    Dim s As String
    Dim ch As Variant
    s = nameIn
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While Len(s) > 0 And (Right$(s, 1) = " " Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "Untitled"
    SanitizeFileName = s
End Function

' --- Helper: strip file extension from a filename ---
Private Function StripExtension(ByVal title As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(title, ".")
    If dotPos > 1 Then
        StripExtension = Left$(title, dotPos - 1)
    Else
        StripExtension = title
    End If
End Function

' --- Helper: perform SaveAs4 + verify ---
Private Function SaveAndCheck( _
    doc As SldWorks.ModelDoc2, _
    path As String, _
    fmtName As String _
) As Boolean
    Dim errs   As Long, warns As Long
    doc.SaveAs4 path, swSaveVerCurrent, swSaveOpSilent, errs, warns
    If errs <> 0 Or Not FileExists(path) Then
        MsgBox fmtName & " export failed." & vbCrLf & _
               "Errors: " & errs & "  Warnings: " & warns, vbCritical
        SaveAndCheck = False
    Else
        SaveAndCheck = True
    End If
End Function

' === Main routine ===
Public Sub ExportStepByConfig()

    Dim swApp        As SldWorks.SldWorks
    Dim swModel      As SldWorks.ModelDoc2
    Dim exportFolder As String
    Dim docType      As Long

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Open a part or assembly before running this macro.", vbCritical
        Exit Sub
    End If

    docType = swModel.GetType
    If docType <> swDocPART And docType <> swDocASSEMBLY Then
        MsgBox "This macro only works on 3D files (.sldprt or .sldasm).", vbCritical
        Exit Sub
    End If

    If Len(swModel.GetPathName) = 0 Then
        MsgBox "Please save the model before running this macro.", vbCritical
        Exit Sub
    End If

    ' -- ask where to export before doing anything else --
    exportFolder = PickExportFolder(DEFAULT_EXPORT_FOLDER)
    If Len(exportFolder) = 0 Then
        ' user cancelled the folder dialog
        Exit Sub
    End If

    If docType = swDocPART Then
        ExportPart swModel, exportFolder
    Else
        ExportAssembly swModel, exportFolder
    End If

End Sub

' --- Build the base name:  <filename> - <active configuration> ---
Private Function ConfigBaseName(swModel As SldWorks.ModelDoc2) As String
    Dim baseNm   As String
    Dim configNm As String
    baseNm = StripExtension(swModel.GetTitle)
    configNm = swModel.ConfigurationManager.ActiveConfiguration.Name
    ConfigBaseName = SanitizeFileName(baseNm & " - " & configNm)
End Function

' === Assembly export (active config -> single STEP) ===
Private Sub ExportAssembly(swModel As SldWorks.ModelDoc2, ByVal exportFolder As String)

    Dim stepP As String
    Dim resp  As VbMsgBoxResult

    stepP = PathJoin(exportFolder, ConfigBaseName(swModel) & ".STEP")

    If FileExists(stepP) Then
        resp = MsgBox("Export file already exists:" & vbCrLf & stepP & vbCrLf & _
                      "Overwrite?", vbQuestion + vbYesNo, "Overwrite file?")
        If resp <> vbYes Then Exit Sub
    End If

    swModel.ClearSelection2 True
    If Not SaveAndCheck(swModel, stepP, "STEP") Then Exit Sub

    MsgBox "Export succeeded:" & vbCrLf & stepP, vbInformation

End Sub

' === Part export (active config -> STEP; optional per-body) ===
Private Sub ExportPart(swModel As SldWorks.ModelDoc2, ByVal exportFolder As String)

    Dim swPart    As SldWorks.PartDoc
    Dim swBody    As SldWorks.Body2
    Dim vBodies   As Variant
    Dim baseNm    As String
    Dim stepP     As String
    Dim resp      As VbMsgBoxResult
    Dim bodyCount As Long
    Dim lb        As Long
    Dim i         As Long

    Set swPart = swModel
    baseNm = ConfigBaseName(swModel)          ' "<filename> - <configuration>"

    ' -- gather solid bodies (in the active configuration) --
    vBodies = swPart.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then
        bodyCount = 0
    Else
        lb = LBound(vBodies)
        bodyCount = UBound(vBodies) - lb + 1
    End If

    If bodyCount = 0 Then
        MsgBox "The part contains no solid bodies to export.", vbCritical
        Exit Sub
    End If

    ' -- ask about per-body export only when there is more than one body --
    Dim perBody As Boolean
    perBody = False
    If bodyCount > 1 Then
        resp = MsgBox("This part has " & bodyCount & " bodies." & vbCrLf & _
                      "Export each body as a separate STEP file?", _
                      vbQuestion + vbYesNo, "Export individual bodies?")
        perBody = (resp = vbYes)
    End If

    If Not perBody Then
        ' -- single STEP of the whole model --
        stepP = PathJoin(exportFolder, baseNm & ".STEP")
        If FileExists(stepP) Then
            resp = MsgBox("Export file already exists:" & vbCrLf & stepP & vbCrLf & _
                          "Overwrite?", vbQuestion + vbYesNo, "Overwrite file?")
            If resp <> vbYes Then Exit Sub
        End If
        swModel.ClearSelection2 True
        If Not SaveAndCheck(swModel, stepP, "STEP") Then Exit Sub
        MsgBox "Export succeeded:" & vbCrLf & stepP, vbInformation
    Else
        ' -- one STEP per body: "<base>-1.STEP", "<base>-2.STEP", ... --
        Dim paths() As String
        Dim existing As String
        ReDim paths(1 To bodyCount)
        existing = ""
        For i = 1 To bodyCount
            paths(i) = PathJoin(exportFolder, baseNm & "-" & i & ".STEP")
            If FileExists(paths(i)) Then existing = existing & vbCrLf & paths(i)
        Next i

        If Len(existing) > 0 Then
            resp = MsgBox("One or more export files already exist:" & existing & vbCrLf & _
                          "Overwrite all?", vbQuestion + vbYesNo, "Overwrite files?")
            If resp <> vbYes Then Exit Sub
        End If

        For i = 1 To bodyCount
            Set swBody = vBodies(lb + i - 1)
            ' selecting a single body makes STEP export only that body
            swModel.ClearSelection2 True
            swBody.Select2 False, Nothing
            If Not SaveAndCheck(swModel, paths(i), "STEP (body " & i & ")") Then
                swModel.ClearSelection2 True
                Exit Sub
            End If
        Next i
        swModel.ClearSelection2 True

        MsgBox "All body exports succeeded (" & bodyCount & " files):" & vbCrLf & _
               paths(1) & " ... " & paths(bodyCount), vbInformation
    End If

End Sub
