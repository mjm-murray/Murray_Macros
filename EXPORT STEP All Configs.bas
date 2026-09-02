Option Explicit

' =====================================================================
'  EXPORT STEP - All Configs
'  ---------------------------------------------------------------
'  Works on 3D files only (.sldprt or .sldasm).
'  Exports EVERY configuration to its own STEP file, named:
'
'      <filename> - <configuration>.STEP
'
'  Output folder is chosen at run time via a folder-browse dialog.
' =====================================================================

' --- Constants ---
Const swDocPART          As Long = 1
Const swDocASSEMBLY      As Long = 2
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
        PathJoin = folder & fileName
    Else
        PathJoin = folder & "\" & fileName
    End If
End Function

' --- Helper: let the user pick an export folder (returns "" if cancelled) ---
Private Function PickExportFolder(ByVal startFolder As String) As String
    Dim shellApp As Object
    Dim folderObj As Object

    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    On Error GoTo 0

    If shellApp Is Nothing Then
        PickExportFolder = startFolder
        Exit Function
    End If

    Set folderObj = shellApp.BrowseForFolder(0&, _
        "Select the folder to export to:", _
        BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE, _
        startFolder)

    If folderObj Is Nothing Then
        PickExportFolder = ""                 ' user cancelled
    Else
        PickExportFolder = folderObj.Self.path
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
        SaveAndCheck = False
    Else
        SaveAndCheck = True
    End If
End Function

' === Main routine ===
Public Sub ExportStepAllConfigs()

    Dim swApp        As SldWorks.SldWorks
    Dim swModel      As SldWorks.ModelDoc2
    Dim exportFolder As String
    Dim docType      As Long

    Dim baseNm       As String
    Dim origConfig   As String
    Dim vConfigs     As Variant
    Dim cfgCount     As Long
    Dim i            As Long

    ' -- init & checks --
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

    ' -- list every configuration --
    vConfigs = swModel.GetConfigurationNames
    If IsEmpty(vConfigs) Then
        MsgBox "No configurations found.", vbCritical
        Exit Sub
    End If
    cfgCount = UBound(vConfigs) - LBound(vConfigs) + 1

    ' -- choose output folder --
    exportFolder = PickExportFolder(DEFAULT_EXPORT_FOLDER)
    If Len(exportFolder) = 0 Then Exit Sub    ' user cancelled

    ' -- remember the currently active config so we can restore it --
    origConfig = swModel.ConfigurationManager.ActiveConfiguration.Name

    ' -- base filename (without config) --
    baseNm = StripExtension(swModel.GetTitle)

    ' -- build target paths and check for existing files --
    Dim paths()  As String
    Dim existing As String
    ReDim paths(0 To cfgCount - 1)
    existing = ""
    For i = 0 To cfgCount - 1
        paths(i) = PathJoin(exportFolder, _
                   SanitizeFileName(baseNm & " - " & CStr(vConfigs(i))) & ".STEP")
        If FileExists(paths(i)) Then existing = existing & vbCrLf & paths(i)
    Next i

    If Len(existing) > 0 Then
        If MsgBox("One or more export files already exist:" & existing & vbCrLf & vbCrLf & _
                  "Overwrite all?", vbQuestion + vbYesNo, "Overwrite files?") <> vbYes Then
            Exit Sub
        End If
    End If

    ' -- loop configs: activate, rebuild, export unconditionally --
    Dim okCount   As Long
    Dim failList  As String
    Dim mismatch  As String
    okCount = 0
    failList = ""
    mismatch = ""

    Dim cfgName As String
    Dim actName As String
    For i = 0 To cfgCount - 1
        cfgName = CStr(vConfigs(i))
        ' Activate the config. Note: ShowConfiguration2 returns False when the
        ' config is already active, so its return value is NOT used as a gate.
        swModel.ShowConfiguration2 cfgName
        swModel.ForceRebuild3 False
        swModel.ClearSelection2 True

        actName = swModel.ConfigurationManager.ActiveConfiguration.Name

        If SaveAndCheck(swModel, paths(i), "STEP") Then
            okCount = okCount + 1
            ' flag (but still count) any case where the active config that was
            ' exported did not match the requested name -- for transparency.
            If StrComp(Trim$(actName), Trim$(cfgName), vbTextCompare) <> 0 Then
                mismatch = mismatch & vbCrLf & cfgName & _
                           "  (exported while active was '" & actName & "')"
            End If
        Else
            failList = failList & vbCrLf & cfgName
        End If
    Next i

    ' -- restore the original active configuration --
    swModel.ShowConfiguration2 origConfig
    swModel.ClearSelection2 True

    ' -- report --
    Dim msg As String
    msg = "Exported " & okCount & " of " & cfgCount & " configurations to:" & vbCrLf & _
          exportFolder
    If Len(mismatch) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Note - active config differed:" & mismatch
    End If
    If Len(failList) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Failed:" & failList
        MsgBox msg, vbExclamation
    ElseIf Len(mismatch) > 0 Then
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If

End Sub
