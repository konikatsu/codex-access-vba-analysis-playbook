Attribute VB_Name = "AccessPlaybookStartupBypass"
Option Compare Database
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function AccessPlaybookGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" _
    (ByVal hWnd As LongPtr, ByRef processId As Long) As Long
#Else
Private Declare Function AccessPlaybookGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" _
    (ByVal hWnd As Long, ByRef processId As Long) As Long
#End If

Private Function AccessPlaybookBooleanJson(ByVal value As Boolean) As String
    If value Then
        AccessPlaybookBooleanJson = "true"
    Else
        AccessPlaybookBooleanJson = "false"
    End If
End Function

Private Function AccessPlaybookIsRunIdSafe(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String

    If Len(value) <> 32 Then Exit Function

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If InStr(1, "0123456789abcdefABCDEF", character, vbBinaryCompare) = 0 Then Exit Function
    Next index

    AccessPlaybookIsRunIdSafe = True
End Function

Public Sub AccessPlaybookAttestStartupBypassIfRequested()
    Dim runId As String
    Dim resultPath As String
    Dim acknowledgePath As String
    Dim expectedDatabasePath As String
    Dim databasePathMatches As Boolean
    Dim openFormCount As Long
#If VBA7 Then
    Dim accessHwnd As LongPtr
#Else
    Dim accessHwnd As Long
#End If
    Dim accessProcessId As Long
    Dim statusText As String
    Dim temporaryPath As String
    Dim fileNumber As Integer
    Dim payload As String
    Dim acknowledgeDeadline As Date

    runId = Environ$("ACCESS_STARTUP_BYPASS_RUN_ID")
    resultPath = Environ$("ACCESS_STARTUP_BYPASS_RESULT_PATH")
    acknowledgePath = Environ$("ACCESS_STARTUP_BYPASS_ACK_PATH")
    expectedDatabasePath = Environ$("ACCESS_STARTUP_BYPASS_EXPECTED_DB_PATH")

    If Len(runId) = 0 And Len(resultPath) = 0 And Len(acknowledgePath) = 0 And Len(expectedDatabasePath) = 0 Then
        Exit Sub
    End If

    On Error GoTo AttestationFailed

    If Not AccessPlaybookIsRunIdSafe(runId) Then Err.Raise 5, , "Invalid startup bypass run ID."
    If Len(resultPath) = 0 Or Len(acknowledgePath) = 0 Or Len(expectedDatabasePath) = 0 Then
        Err.Raise 5, , "Incomplete startup bypass attestation environment."
    End If

    databasePathMatches = (StrComp(CurrentProject.FullName, expectedDatabasePath, vbTextCompare) = 0)
    openFormCount = Forms.Count
    accessHwnd = Application.hWndAccessApp
    Call AccessPlaybookGetWindowThreadProcessId(accessHwnd, accessProcessId)

    If databasePathMatches And openFormCount = 0 And accessHwnd <> 0 And accessProcessId > 0 Then
        statusText = "PASS"
    Else
        statusText = "FAIL"
    End If

    payload = "{""schema_version"":1,""run_id"":""" & runId & _
        """,""command"":""SKIP_AUTOEXEC"",""status"":""" & statusText & _
        """,""database_path_match"":" & AccessPlaybookBooleanJson(databasePathMatches) & _
        ",""forms_count"":" & CStr(openFormCount) & _
        ",""hwnd_access_app"":" & CStr(accessHwnd) & _
        ",""process_id"":" & CStr(accessProcessId) & "}"

    temporaryPath = resultPath & "." & runId & ".tmp"
    fileNumber = FreeFile
    Open temporaryPath For Output Access Write Lock Write As #fileNumber
    Print #fileNumber, payload
    Close #fileNumber
    fileNumber = 0
    Name temporaryPath As resultPath

    acknowledgeDeadline = DateAdd("s", 10, Now)
    Do While Len(Dir$(acknowledgePath, vbNormal Or vbHidden Or vbSystem)) = 0
        DoEvents
        If Now >= acknowledgeDeadline Then Exit Do
    Loop

    Application.Quit acQuitSaveNone
    Exit Sub

AttestationFailed:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    Application.Quit acQuitSaveNone
End Sub
