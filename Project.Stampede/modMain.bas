Attribute VB_Name = "modMain"
Option Explicit

Private oStamp  As cStampede
Dim frm         As frmProgress

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function PathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Sub Main()

    If Command$ <> "" Then
        Set frm = New frmProgress
        parseCommandStr Command$, frm
        Unload frm
        End
    Else
        frmAbout.Show
    End If
    
End Sub
Public Sub parseCommandStr(Command As String, ByRef frm As frmProgress)
On Error GoTo err_Handler
Dim Str         As String
Dim fAry()      As String
Dim i           As Long

        If Len(Trim(Command)) = 0 Then Exit Sub
        'Here we check the Command String to see if it
        'is a list of files

        If IsFileArray(Command$) Then
                fAry = BuildFileArray(Command$)
                For i = 0 To UBound(fAry)
                    parseCommandStr fAry(i), frm
                Next i
        End If
        
        Set oStamp = New cStampede
        Str = Replace(Command, Chr(34), "")
        
        If CBool(PathIsDirectory(Str)) Then
            If CBool(PathIsDirectoryEmpty(Str)) Then
                'DELETE DIRECTORY STRUCTURE ONLY
                RemoveDirectory Str
            Else
                'RECURSIVE FILE DELETION HERE
                Dim oSrch As New cFileSearch
                
                    oSrch.FileSearch Str & "*.*"
                    For i = 1 To oSrch.FileCount
                        'Delete files using the same Proc we are in
                        parseCommandStr oSrch.FileItem(i).AFilePath, frm
                    Next i
                    For i = 1 To oSrch.DirCount
                        'Delete dirs using the same proc we are in
                        parseCommandStr oSrch.DirItem(i).DirPath, frm
                    Next i
                    'Delete the root folder
                    parseCommandStr Str, frm
                    Set oSrch = Nothing
                
            End If
        ElseIf CBool(PathFileExists(Str)) Then
            'DELETE THE FILE WITH AN 8x OVERWRITE
            If Not frm.Visible Then frm.Show
            frm.sb1.Panels(2).Text = Str
            oStamp.STAMPEDE Str, frm.pb1
            frm.sb1.Panels(1).Text = (oStamp.file_OverWriteTime / 1000) & " Secs"
        End If
        
        Exit Sub
        
err_Handler:
    MsgBox "Main Module " & Err.Number & ":" & Err.Description & ":" & Err.LastDllError
    Resume Next
        
        
End Sub
Private Function IsFileArray(sCmd As String) As Boolean
Dim i           As Long
Dim iOcc        As Long

    'This is an cheat test, a file name cannot contain the character ':' so
    'if it is a file array then it must have more than 1 ':' appear in the line
    'ie: c:\Test.txt c:\wedlock.exe, you may be thinking we should be looking
    ' at the spaces but most file systems now allow for names like 'c:\My File.txt'
    
    For i = 1 To Len(sCmd)
        If Mid(sCmd, i, 1) = ":" Then iOcc = iOcc + 1
    Next i
    
    If iOcc > 1 Then IsFileArray = True

End Function
Private Function BuildFileArray(sCmd As String) As String()
Dim tmpString As String
Dim index As Integer
Dim iEnd As Integer
    
    tmpString = ""
    sCmd = Replace(sCmd, Chr(34), "")
    Do While Len(sCmd) > 0
    index = InStr(1, sCmd, ":", vbBinaryCompare)
    If index > 0 Then
        iEnd = InStr(index + 1, sCmd, ":", vbBinaryCompare) - 1
        If iEnd > 0 Then
        tmpString = tmpString & Trim(Mid(sCmd, 1, iEnd - 1)) & ","
        sCmd = Right(sCmd, Len(sCmd) - Len(Mid(sCmd, 1, iEnd - 1)))
        Else
            tmpString = tmpString & sCmd
            sCmd = ""
        End If
    End If
    Loop
    If Right(tmpString, 1) = "," Then MsgBox (tmpString): tmpString = Left(tmpString, Len(tmpString) - 1)

    BuildFileArray = Split(tmpString, ",")
    
      
End Function
