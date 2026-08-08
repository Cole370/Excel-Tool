Attribute VB_Name = "Sandbox"
Option Explicit

'==============================================================================
' Sandbox.bas  -  Build step 1: the filesystem path guard.
'
' Claude's filesystem reach is restricted to SANDBOX_ROOT and everything inside
' it. No other path may ever be read or written. This module is the single
' choke point that enforces that rule. Every filesystem call in the tool must
' pass its path through IsInSandbox() before touching disk.
'
' Design note (see spec section 2.1): we RESOLVE the path first (collapsing
' "." and ".." and forward slashes), then COMPARE against the root WITH A
' MANDATORY TRAILING SEPARATOR. Without the trailing separator,
' "...\INDEX_backup\secrets.xlsx" would sneak past a naive prefix test against
' "...\INDEX". That is the single easiest way to silently break the sandbox,
' so it is tested explicitly in Test_Sandbox below.
'==============================================================================

Private Const SANDBOX_ROOT As String = "C:\Users\colet\Documents\INDEX"

' Test tally (used only by Test_Sandbox)
Private mPass As Long
Private mTotal As Long

'------------------------------------------------------------------------------
' IsInSandbox - True only if 'candidate' resolves to a path inside SANDBOX_ROOT.
'------------------------------------------------------------------------------
Public Function IsInSandbox(ByVal candidate As String) As Boolean
    Dim fso As Object, resolved As String, root As String
    IsInSandbox = False

    candidate = Trim$(candidate)
    If Len(candidate) = 0 Then Exit Function

    ' Normalise forward slashes to backslashes so "/" cannot dodge the checks.
    candidate = Replace(candidate, "/", "\")

    ' Reject UNC network paths (\\server\share\...).
    If Left$(candidate, 2) = "\\" Then Exit Function
    ' Reject alternate-data-stream forms (file.xlsx::$DATA).
    If InStr(candidate, "::") > 0 Then Exit Function
    ' Reject any colon beyond the drive-letter colon at position 2
    ' (blocks single-colon ADS like  ...\a.xlsx:Zone.Identifier ).
    If InStr(3, candidate, ":") > 0 Then Exit Function

    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    resolved = fso.GetAbsolutePathName(candidate)   ' collapses . and ..
    root = fso.GetAbsolutePathName(SANDBOX_ROOT)

    ' TRAILING SEPARATOR IS MANDATORY - do not remove.
    If Right$(root, 1) <> "\" Then root = root & "\"

    IsInSandbox = (LCase$(Left$(resolved & "\", Len(root))) = LCase$(root))
    Exit Function
Fail:
    IsInSandbox = False
End Function

'------------------------------------------------------------------------------
' IsAllowedExtension - True only for the whitelisted spreadsheet/text/pdf types.
'------------------------------------------------------------------------------
Public Function IsAllowedExtension(ByVal candidate As String) As Boolean
    Dim ext As String, allowed As Variant, i As Long, dotPos As Long
    IsAllowedExtension = False

    dotPos = InStrRev(candidate, ".")
    If dotPos = 0 Then Exit Function
    ext = LCase$(Mid$(candidate, dotPos))

    allowed = Array(".xlsx", ".xlsm", ".xlsb", ".xls", ".csv", ".tsv", ".txt", ".pdf")
    For i = LBound(allowed) To UBound(allowed)
        If ext = allowed(i) Then
            IsAllowedExtension = True
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' IsReparsePoint - True if a folder is a junction/symlink (reparse point).
' Used later by the indexer tree-walk to skip folders that could point outside
' the sandbox. Live-tested when the indexer is built (step 4).
'------------------------------------------------------------------------------
Public Function IsReparsePoint(ByVal folderPath As String) As Boolean
    Dim fso As Object, f As Object
    IsReparsePoint = False
    On Error GoTo done
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    Set f = fso.GetFolder(folderPath)
    IsReparsePoint = ((f.Attributes And 1024) = 1024)   ' FILE_ATTRIBUTE_REPARSE_POINT
done:
End Function

'==============================================================================
' Test_Sandbox - run this (F5) and read the Immediate Window (Ctrl+G).
' Expected final line: === RESULT: ALL 23 TESTS PASSED ===
'==============================================================================
Public Sub Test_Sandbox()
    mPass = 0
    mTotal = 0
    Debug.Print "=== SANDBOX PATH GUARD TESTS ==="
    Debug.Print "Root = " & SANDBOX_ROOT

    ' --- IsInSandbox: paths that SHOULD be allowed (inside) ---
    Assert "inside: normal file in root", IsInSandbox("C:\Users\colet\Documents\INDEX\Model.xlsx"), True
    Assert "inside: the root folder itself", IsInSandbox("C:\Users\colet\Documents\INDEX"), True
    Assert "inside: nested subfolder file", IsInSandbox("C:\Users\colet\Documents\INDEX\sub\deep\file.xlsx"), True
    Assert "inside: traversal that stays inside", IsInSandbox("C:\Users\colet\Documents\INDEX\..\INDEX\ok.xlsx"), True
    Assert "inside: case-insensitive path", IsInSandbox("c:\users\COLET\documents\index\Model.xlsx"), True

    ' --- IsInSandbox: paths that MUST be blocked (outside / malicious) ---
    Assert "block: INDEX_backup prefix trap", IsInSandbox("C:\Users\colet\Documents\INDEX_backup\secrets.xlsx"), False
    Assert "block: INDEXsecret prefix trap", IsInSandbox("C:\Users\colet\Documents\INDEXsecret\x.xlsx"), False
    Assert "block: ..\..\ escapes upward", IsInSandbox("C:\Users\colet\Documents\INDEX\..\..\secret.xlsx"), False
    Assert "block: deep traversal escape", IsInSandbox("C:\Users\colet\Documents\INDEX\sub\..\..\..\Windows\x.xlsx"), False
    Assert "block: UNC network path", IsInSandbox("\\server\share\INDEX\x.xlsx"), False
    Assert "block: alternate data stream (::)", IsInSandbox("C:\Users\colet\Documents\INDEX\file.xlsx::$DATA"), False
    Assert "block: single-colon ADS", IsInSandbox("C:\Users\colet\Documents\INDEX\a.xlsx:Zone.Identifier"), False
    Assert "block: empty string", IsInSandbox(""), False
    Assert "block: whitespace only", IsInSandbox("     "), False

    ' --- IsAllowedExtension ---
    Assert "ext: .xlsx allowed", IsAllowedExtension("Model.xlsx"), True
    Assert "ext: .csv allowed", IsAllowedExtension("data.csv"), True
    Assert "ext: .txt allowed", IsAllowedExtension("notes.txt"), True
    Assert "ext: .pdf allowed", IsAllowedExtension("report.pdf"), True
    Assert "ext: uppercase .XLSX allowed", IsAllowedExtension("MODEL.XLSX"), True
    Assert "ext: .exe blocked", IsAllowedExtension("malware.exe"), False
    Assert "ext: .vbs blocked", IsAllowedExtension("script.vbs"), False
    Assert "ext: .zip blocked", IsAllowedExtension("archive.zip"), False
    Assert "ext: no extension blocked", IsAllowedExtension("noextension"), False

    Debug.Print "-------------------------------------"
    If mPass = mTotal Then
        Debug.Print "=== RESULT: ALL " & mTotal & " TESTS PASSED ==="
    Else
        Debug.Print "=== RESULT: " & mPass & "/" & mTotal & " passed - " & (mTotal - mPass) & " FAILED ==="
    End If
End Sub

Private Sub Assert(ByVal desc As String, ByVal got As Boolean, ByVal expected As Boolean)
    mTotal = mTotal + 1
    If got = expected Then
        mPass = mPass + 1
        Debug.Print "[PASS] " & desc
    Else
        Debug.Print "[FAIL] " & desc & "  (expected " & expected & ", got " & got & ")"
    End If
End Sub
