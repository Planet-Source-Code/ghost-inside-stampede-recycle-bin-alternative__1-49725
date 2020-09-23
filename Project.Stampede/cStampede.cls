VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStampede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##########################################################################
'# AUTHOR: John Bridle                                                    #
'#                                                                        #
'# PURPOSE: The stampede class overwrites a file 8 times at super speeds  #
'#          and then flushes the file buffer before deleteing a file.     #
'#          In many other classes of this nature the File buffer flush    #
'#          has not been actioned correctly.  Windows is set up to cache  #
'#          file changes if a file is marked for delete before the cache  #
'#          is flushed overwrites do not take place.                      #
'#                                                                        #
'# CREATION DATE: 5/11/2003                                               #
'# LAST UPDATED: 5/11/2003                                                #
'# COPYRIGHT RAINYDAYZ.COM 2003 'Not to be replicated without prior       #
'#                                permission from John Bridle'            #
'# EMAIL: jbridle@rainydayz.com                                           #
'#                                                                        #
'##########################################################################
        
'FILE READ AND WRITE CONSTANTS
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_BEGIN = 0
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_NEW = 1
Private Const OPEN_EXISTING = 3
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const INVALID_HANDLE_VALUE = -1&

'LOCALISED CONSTANTS
Private Const ChunkSize     As Long = 16384
Private Const BAD_RET_VALUE = 0&

'Private Declared variables
Private tOWTime             As Long
Private tAryTime            As Long

'API FUNCTIONS DECLARED
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Sub PathQuoteSpaces Lib "shlwapi.dll" Alias "PathQuoteSpacesA" (ByVal lpsz As String)
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Const LANG_NEUTRAL = &H0
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'PUBLIC EVENTS
Public Event OnError(ErrNum As Long, Errdesc As String)
Public Sub STAMPEDE(path As String, Optional ByRef pb As ProgressBar)
On Error GoTo err_Handler
Dim X               As Long
Dim Y               As Long
Dim lFileSz         As Long
Dim HwFile          As Long
Dim fBytes()        As Byte
Dim low             As Long
Dim lChnk           As Long
Dim lDiff           As Long
Dim lret            As Long
Dim buffer          As String
Dim HavePb          As Boolean
    
    If Not IsMissing(pb) Then HavePb = True: pb.Value = 0
    'Monitor time taken to create array
    tAryTime = timeGetTime
    buffer = Space(200)
    'Remove any attributes that may halt the process
    SetFileAttributes path, FILE_ATTRIBUTE_NORMAL
    'Prepare the file for sequential access and collect a Handle for the required file
    HwFile = CreateFile(path, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    'Check the file handle is valid before proceeding
    If HwFile = INVALID_HANDLE_VALUE Then
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Err.LastDllError, LANG_NEUTRAL, buffer, 200, ByVal 0&
        RaiseEvent OnError(Err.LastDllError, buffer): Exit Sub
        MsgBox buffer
    End If
    'Get the file size using the file handle
    lFileSz = GetFileSize(HwFile, 0)
    'If the file contains no bytes ...nothing to overwrite
    If lFileSz = 0 Then GoTo EndProcess
    'Calculate the number of chunks the overwrite will have to loop through
    'we write with chunks for lower memory overhead for large file sizes
    If lFileSz > ChunkSize Then lChnk = RoundDown(lFileSz / ChunkSize, 1)
    'The remainder of bytes left to tag as seperate process at the end
    lDiff = lFileSz - (lChnk * ChunkSize)
    If lDiff > 0 And HavePb Then pb.Max = lChnk + 1 Else pb.Max = lChnk
    'Create a byte array and fill it with characters
    buildBytes ChunkSize, fBytes
    
    'Stop recording byte array time
    tAryTime = timeGetTime - tAryTime
    'Start recording time to overwrite and delete file
    tOWTime = timeGetTime
    
    Screen.MousePointer = vbHourglass
    'Start Chunk style overwrite of file utilising the Sequential access flag buffers created
    'when preparing the file for opening earlier.  Repeat this step for eight overwrites...not
    'very likely that anyone will be able to recover a file from eight writes but I would love to be defied
    For X = 1 To 8
        For Y = 0 To lChnk
            'FILE IS NOW BEING SERIOUSLY STAMPEDED (IS THAT A REAL WORD LOL)
            'First we move our pointer to the correct position in the file
            low = SetFilePointer(HwFile, CLng(Y * ChunkSize), 0, FILE_BEGIN)
            'Now we write to over the file
            WriteFile HwFile, fBytes(1), UBound(fBytes), lret, ByVal 0&
            pb.Value = Y
            DoEvents
        Next Y
    Next X
    
    'Build a new byte array for the Chunks overhang created when segmenting file size
    If lDiff > 0 Then buildBytes lDiff, fBytes
    
    'Now repeat the above process for the file overhang
    For X = 1 To 8
        low = SetFilePointer(HwFile, CLng(lFileSz - lDiff), 0, FILE_BEGIN)
        WriteFile HwFile, fBytes(1), UBound(fBytes), lret, ByVal 0&
        pb.Value = lChnk
        DoEvents
    Next X
    
    'Before deleteing the file, flushing the Filebuffers causes the changes to be commited
    'to hard drive
    FlushFileBuffers HwFile
    'Close the file handle and release the file
    'from our process
EndProcess:
    CloseHandle HwFile

    'Finish recording the time taken for overwrite
    tOWTime = timeGetTime - tOWTime
    'MsgBox "Part 1 Array Spoke " & CStr((tAryTime) / 1000) & " secs to complete" & vbCrLf & _
            '"Part 2 File Write " & CStr((tOWTime) / 1000) & " secs to complete"
    'Delete the file thus destroying it for good
    DeleteFile (path)
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err_Handler:
    MsgBox "Stampede Class " & Err.Number & ":" & Err.Description
    Resume Next
    
    
End Sub
Private Sub buildBytes(lsize As Long, ByRef fBytes() As Byte)
Dim lB      As Long

    ReDim fBytes(1 To lsize) As Byte
    For lB = 1 To UBound(fBytes)
        fBytes(lB) = Rnd * 255
    Next lB
    
End Sub
Public Property Get Array_buildTime() As Long
    Array_buildTime = tAryTime
End Property
Public Property Get file_OverWriteTime() As Long
    file_OverWriteTime = tOWTime
End Property
Function RoundDown(ByVal Value As Long, Round As Long)
       RoundDown = Value - (Value Mod Round)
End Function


