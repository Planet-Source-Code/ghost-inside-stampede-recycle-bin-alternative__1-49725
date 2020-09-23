VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecursive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fpCol As Collection
Private drCol As Collection


Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1&
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const BAD_FILE_RET = 0

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type typeFileSection
    sFilename As String
    sFileExt As String
    sFilePath As String
End Type


Private Type WIN32_FIND_DATA
    W32FileData As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Private oFile As cFiles
Private oDir As cDirs
Private lFileCount As Long
Private lDirCount As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Sub IsNewSearch(FilesDeep As Long)

    If FilesDeep = 1 Then Set fpCol = New Collection: Set drCol = NewCollection
    
End Sub
Public Sub FileSearch(sDirectory As String)
Dim HwFile As Long
Dim HwNFile As Long
Dim W32FileData As WIN32_FIND_DATA
Static FilesDeep As Long

    FilesDeep = FilesDeep + 1
    IsNewSearch FilesDeep

    HwFile = FindFirstFile(sDirectory, W32FileData)
    NewEntryFound sDirectory, W32FileData
    
    If HwFile <> INVALID_HANDLE_VALUE Then
        Do While HwNFile > BAD_FILE_RET
            HwNFile = FindNextFile(HwFile, W32FileData)
            NewEntryFound sDirectory, W32FileData
        Loop
        
        FindClose HwFile
    End If
    
    thisLevel = thisLevel - 1
    
End Sub
Private Sub NewEntryFound(sDirectory As String, W32FileData As WIN32_FIND_DATA)

        If W32FileData.W32FileData <> FILE_ATTRIBUTE_DIRECTORY Then
            AddFile sDirectory, W32FileData
            lFileCount = lFileCount + 1
        ElseIf W32FileData.W32FileData = FILE_ATTRIBUTE_DIRECTORY Then
            AddDir sDirectory, W32FileData
        End If
    
End Sub
Private Sub AddFile(sDirectory As String, W32FileData As WIN32_FIND_DATA)

        Set oFile = New cFiles
        oFile.AFilePath = FileParts(sDirectory).sFilePath & fTrimString32(W32FileData.cFileName)
        fpCol.Add oFile
        
End Sub
Private Function AddDir(sDirectory As String, W32FileData As WIN32_FIND_DATA) As String

        Set oDir = New cDirs
        oDir.DirPath = FileParts(sDirectory).sFilePath & fTrimString32(W32FileData.cFileName) & "\"
        drCol.Add oDir
        AddDir = oDir.DirPath
            
End Function
Private Function fTrimString32(ByVal sMyString As String) As String
Dim isStart As Integer

    'This Function is used internally by this class alone
    'it trims strings that have been pre-buffered and leaves
    'remaining the windows 32bit file name, 16bit file names
    'are returned in the file name string but are lost in this process
    isStart = InStr(1, sMyString, Chr(0))
    
    If isStart = 0 Then Exit Function
    
    fTrimString32 = Left(sMyString, isStart - 1)

End Function
Private Function FileParts(ByVal sFilePath As String) As typeFileSection
Dim iDot1, iSlash1 As Integer
Dim tFS As typeFileSection

    iDot1 = InStrRev(sFilePath, ".")
    iSlash1 = InStrRev(sFilePath, "\")
    
    If iSlash1 > 0 Then
        tFS.sFilePath = Left(sFilePath, iSlash1)
    End If
    
    If iDot1 > 0 And iDot1 > iSlash1 Then
        tFS.sFileExt = Mid(sFilePath, iDot1 + 1)
        tFS.sFilename = Mid(sFilePath, iSlash1 + 1, iDot1 - iSlash1 - 1)
    End If
    
    fFileParts = tFS
    
End Function
Public Property Get FileItem(vntIndexKey As Variant) As cFiles
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
  Set FileItem = fpCol(vntIndexKey)
  
End Property
Public Property Get FileCount() As Long
    FileCount = lFileCount
End Property
