Attribute VB_Name = "GsTools_analysisinfo"
Option Compare Database
Option Explicit

' Export information useful for static analysis of the current Access database.
Public Sub ExportAnalysisInfo()
    Dim outputRoot As String
    Dim outputDir As String
    Dim latestDir As String
    Dim objectsDir As String
    Dim schemaDir As String
    Dim logPath As String
    Dim currentDat As Object
    Dim currentProj As Object
    Dim fso As Object
    Dim logNo As Integer

    On Error GoTo FatalError

    Set fso = CreateObject("Scripting.FileSystemObject")

    outputRoot = CurrentProject.Path & "\Defines" & CurrentProject.Name
    outputDir = outputRoot & "\Exports\" & Format$(Now, "yyyymmdd_hhnnss")
    latestDir = outputRoot & "\Latest"
    objectsDir = outputDir & "\Objects"
    schemaDir = outputDir & "\Schema"

    EnsureFolder fso, objectsDir & "\Forms"
    EnsureFolder fso, objectsDir & "\Reports"
    EnsureFolder fso, objectsDir & "\Macros"
    EnsureFolder fso, objectsDir & "\Modules"
    EnsureFolder fso, objectsDir & "\Queries"
    EnsureFolder fso, schemaDir

    logPath = outputDir & "\ExportAnalysisInfo.log"
    logNo = FreeFile
    Open logPath For Output As #logNo

    LogLine logNo, "Started: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    LogLine logNo, "Database: " & CurrentDb.Name
    LogLine logNo, "OutputDir: " & outputDir

    Set currentDat = Application.CurrentData
    Set currentProj = Application.CurrentProject

    ExportObjectType acForm, currentProj.AllForms, objectsDir & "\Forms", ".frm", logNo
    ExportObjectType acReport, currentProj.AllReports, objectsDir & "\Reports", ".rpt", logNo
    ExportObjectType acMacro, currentProj.AllMacros, objectsDir & "\Macros", ".mcr", logNo
    ExportObjectType acModule, currentProj.AllModules, objectsDir & "\Modules", ".mdl", logNo
    ExportObjectType acQuery, currentDat.AllQueries, objectsDir & "\Queries", ".qry", logNo

    ExportTableDefinitions schemaDir & "\Tables.txt", logNo
    ExportRelations schemaDir & "\Relations.txt", logNo
    ExportReferences schemaDir & "\References.txt", logNo
    ExportDatabaseProperties schemaDir & "\DatabaseProperties.txt", logNo

    LogLine logNo, "Finished: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #logNo
    UpdateLatestFolder fso, outputDir, latestDir, 0
    Debug.Print "ExportAnalysisInfo finished: " & outputDir
    Exit Sub

FatalError:
    If logNo <> 0 Then
        LogLine logNo, "FATAL " & Err.Number & ": " & Err.Description
        On Error Resume Next
        Close #logNo
    Else
        Debug.Print "FATAL " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub UpdateLatestFolder(ByVal fso As Object, ByVal sourceDir As String, _
        ByVal latestDir As String, ByVal logNo As Integer)

    On Error GoTo ErrHandler

    EnsureFolder fso, latestDir
    ClearFolderContents fso, latestDir
    CopyFolderContents fso, sourceDir, latestDir
    LogLine logNo, "Updated Latest: " & latestDir
    Exit Sub

ErrHandler:
    LogLine logNo, "ERROR UpdateLatest " & Err.Number & ": " & Err.Description
End Sub

Private Sub ExportObjectType(ByVal objType As Integer, _
        ByVal objCollection As Variant, ByVal outputDir As String, _
        ByVal ext As String, ByVal logNo As Integer)

    Dim obj As Variant
    Dim filePath As String

    For Each obj In objCollection
        If ShouldExportObject(CStr(obj.Name)) Then
            filePath = outputDir & "\" & SafeFileName(CStr(obj.Name)) & ext
            Err.Clear
            On Error Resume Next
            SaveAsText objType, obj.Name, filePath
            If Err.Number <> 0 Then
                LogLine logNo, "ERROR SaveAsText " & obj.Name & " -> " & Err.Number & ": " & Err.Description
                Err.Clear
            Else
                LogLine logNo, "Save " & obj.Name
            End If
            On Error GoTo 0
        Else
            LogLine logNo, "Skip " & obj.Name
        End If
    Next
End Sub

Private Function ShouldExportObject(ByVal objectName As String) As Boolean
    ShouldExportObject = (Left$(objectName, 1) <> "~" And Left$(objectName, 4) <> "MSys")
End Function

Private Sub ExportTableDefinitions(ByVal filePath As String, ByVal logNo As Integer)
    Dim db As Object
    Dim tdf As Object
    Dim fld As Object
    Dim idx As Object
    Dim idxFld As Object
    Dim fileNo As Integer

    On Error GoTo ErrHandler
    Set db = CurrentDb
    fileNo = FreeFile
    Open filePath For Output As #fileNo

    For Each tdf In db.TableDefs
        If ShouldExportObject(CStr(tdf.Name)) Then
            Print #fileNo, "TABLE " & tdf.Name
            Print #fileNo, "  Attributes: " & tdf.Attributes
            If Len(tdf.Connect & "") > 0 Then Print #fileNo, "  Connect: " & tdf.Connect
            If Len(tdf.SourceTableName & "") > 0 Then Print #fileNo, "  SourceTableName: " & tdf.SourceTableName
            Print #fileNo, "  Fields:"
            For Each fld In tdf.Fields
                Print #fileNo, "    " & fld.Name & _
                    " Type=" & fld.Type & _
                    " Size=" & fld.Size & _
                    " Required=" & SafePropertyValue(fld, "Required") & _
                    " Default=" & SafePropertyValue(fld, "DefaultValue") & _
                    " ValidationRule=" & SafePropertyValue(fld, "ValidationRule")
            Next
            Print #fileNo, "  Indexes:"
            For Each idx In tdf.Indexes
                Print #fileNo, "    " & idx.Name & _
                    " Primary=" & idx.Primary & _
                    " Unique=" & idx.Unique & _
                    " Required=" & idx.Required
                For Each idxFld In idx.Fields
                    Print #fileNo, "      " & idxFld.Name
                Next
            Next
            Print #fileNo, ""
        End If
    Next

    Close #fileNo
    LogLine logNo, "Save Schema\Tables.txt"
    Exit Sub

ErrHandler:
    LogLine logNo, "ERROR ExportTableDefinitions " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #fileNo
End Sub

Private Sub ExportRelations(ByVal filePath As String, ByVal logNo As Integer)
    Dim db As Object
    Dim rel As Object
    Dim fld As Object
    Dim fileNo As Integer

    On Error GoTo ErrHandler
    Set db = CurrentDb
    fileNo = FreeFile
    Open filePath For Output As #fileNo

    For Each rel In db.Relations
        Print #fileNo, "RELATION " & rel.Name
        Print #fileNo, "  Table: " & rel.Table
        Print #fileNo, "  ForeignTable: " & rel.ForeignTable
        Print #fileNo, "  Attributes: " & rel.Attributes
        For Each fld In rel.Fields
            Print #fileNo, "    " & fld.Name & " -> " & fld.ForeignName
        Next
        Print #fileNo, ""
    Next

    Close #fileNo
    LogLine logNo, "Save Schema\Relations.txt"
    Exit Sub

ErrHandler:
    LogLine logNo, "ERROR ExportRelations " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #fileNo
End Sub

Private Sub ExportReferences(ByVal filePath As String, ByVal logNo As Integer)
    Dim ref As Object
    Dim fileNo As Integer

    On Error GoTo ErrHandler
    fileNo = FreeFile
    Open filePath For Output As #fileNo

    For Each ref In Application.References
        Print #fileNo, "REFERENCE " & ref.Name
        Print #fileNo, "  FullPath: " & ref.FullPath
        Print #fileNo, "  Guid: " & ref.Guid
        Print #fileNo, "  Version: " & ref.Major & "." & ref.Minor
        Print #fileNo, "  IsBroken: " & ref.IsBroken
        Print #fileNo, ""
    Next

    Close #fileNo
    LogLine logNo, "Save Schema\References.txt"
    Exit Sub

ErrHandler:
    LogLine logNo, "ERROR ExportReferences " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #fileNo
End Sub

Private Sub ExportDatabaseProperties(ByVal filePath As String, ByVal logNo As Integer)
    Dim db As Object
    Dim prop As Object
    Dim fileNo As Integer

    On Error GoTo ErrHandler
    Set db = CurrentDb
    fileNo = FreeFile
    Open filePath For Output As #fileNo

    Print #fileNo, "Name: " & db.Name
    Print #fileNo, "CurrentProject.Name: " & CurrentProject.Name
    Print #fileNo, "CurrentProject.Path: " & CurrentProject.Path
    Print #fileNo, ""
    Print #fileNo, "DATABASE PROPERTIES"
    For Each prop In db.Properties
        Print #fileNo, prop.Name & ": " & SafePropertyValue(db, prop.Name)
    Next

    Close #fileNo
    LogLine logNo, "Save Schema\DatabaseProperties.txt"
    Exit Sub

ErrHandler:
    LogLine logNo, "ERROR ExportDatabaseProperties " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #fileNo
End Sub

Private Sub EnsureFolder(ByVal fso As Object, ByVal folderPath As String)
    Dim parentPath As String

    If fso.FolderExists(folderPath) Then Exit Sub
    parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 Then EnsureFolder fso, parentPath
    fso.CreateFolder folderPath
End Sub

Private Sub ClearFolderContents(ByVal fso As Object, ByVal folderPath As String)
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object

    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.Files
        file.Delete True
    Next

    For Each subFolder In folder.SubFolders
        subFolder.Delete True
    Next
End Sub

Private Sub CopyFolderContents(ByVal fso As Object, ByVal sourceDir As String, _
        ByVal targetDir As String)

    If fso.FolderExists(sourceDir) Then
        fso.CopyFile sourceDir & "\*.*", targetDir & "\", True
        fso.CopyFolder sourceDir & "\*", targetDir & "\", True
    End If
End Sub

Private Function SafeFileName(ByVal value As String) As String
    Dim badChars As Variant
    Dim reservedNames As Variant
    Dim i As Long
    Dim item As Variant

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        value = Replace$(value, badChars(i), "_")
    Next

    Do While Len(value) > 0 And (Right$(value, 1) = "." Or Right$(value, 1) = " ")
        value = Left$(value, Len(value) - 1)
    Loop

    If Left$(value, 1) = "." Then value = "_" & Mid$(value, 2)

    reservedNames = Array("CON", "PRN", "AUX", "NUL", _
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
    For Each item In reservedNames
        If UCase$(value) = CStr(item) Then value = value & "_"
    Next

    If Len(value) > 120 Then value = Left$(value, 120)
    If Len(value) = 0 Then value = "_"
    SafeFileName = value
End Function

Private Function SafePropertyValue(ByVal obj As Object, ByVal propertyName As String) As String
    On Error GoTo ErrHandler
    SafePropertyValue = CStr(obj.Properties(propertyName).Value)
    Exit Function

ErrHandler:
    SafePropertyValue = ""
End Function

Private Sub LogLine(ByVal logNo As Integer, ByVal message As String)
    Debug.Print message
    If logNo <> 0 Then Print #logNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & message
End Sub
