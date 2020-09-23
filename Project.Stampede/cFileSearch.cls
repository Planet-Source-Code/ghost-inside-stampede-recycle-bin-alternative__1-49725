VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##########################################################################
'# AUTHOR: John Bridle                                                    #
'#                                                                        #
'# PURPOSE: Very fast recursive file/folder search to be used in          #
'#          conjunction with cDirs & cfiles class  and modFileStuff       #
'#                                                                        #
'# CREATION DATE: 5/11/2003                                               #
'# LAST UPDATED: 5/11/2003                                                #
'# COPYRIGHT RAINYDAYZ.COM 2003 'Not to be replicated without prior       #
'#                                permission from John Bridle'            #
'# EMAIL: jbridle@rainydayz.com                                           #
'#                                                                        #
'##########################################################################
        
'Private Declared Variables
Private fpCol As Collection
Private drCol As Collection
Private oFile As cFiles
Private oDir As cDirs
Private lFileCount As Long
Private lDirCount As Long

'Private Declared API functions
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Sub IsNewSearch(FilesDeep As Long)

    'Monitors how deep a recursive search is, if it equals 1 then the collections can be
    'restarted as a new search is taking place
    If FilesDeep = 1 Then Set fpCol = New Collection: Set drCol = New Collection
    
End Sub
Public Sub FileSearch(sDirectory As String)
Dim HwFile As Long
Dim HwNFile As Long
Dim W32FileData As WIN32_FIND_DATA
Static FilesDeep As Long

    'Increases recursive folder depth
    FilesDeep = FilesDeep + 1
    IsNewSearch FilesDeep

    'locate first file and record its handle
    HwFile = FindFirstFile(sDirectory, W32FileData)
    
    'Check to see if we have a valid handle
    If HwFile <> INVALID_HANDLE_VALUE Then
        'Add the fist file to the collection
        NewEntryFound sDirectory, W32FileData
        'Sets return value before entering loop
        HwNFile = 1
        Do While HwNFile > BAD_FILE_RET
            'Get next file
            HwNFile = FindNextFile(HwFile, W32FileData)
            'If return value good, add to collection
            If HwNFile <> BAD_FILE_RET Then NewEntryFound sDirectory, W32FileData
        Loop
        'Close file handle
        FindClose HwFile
    End If
    
    'decreases recursive folder depth
    FilesDeep = FilesDeep - 1
    
End Sub
Private Sub NewEntryFound(sDirectory As String, W32FileData As WIN32_FIND_DATA)

        'Analyse file type...is file or directory and add it to the appropriate
        'collection
        If W32FileData.W32FileData <> FILE_ATTRIBUTE_DIRECTORY Then
            AddFile sDirectory, W32FileData
            lFileCount = lFileCount + 1
        ElseIf W32FileData.W32FileData = FILE_ATTRIBUTE_DIRECTORY Then
            If fTrimString32(W32FileData.cFileName) <> "." And fTrimString32(W32FileData.cFileName) <> ".." Then
                'Here we start the recursive search part
                FileSearch AddDir(sDirectory, W32FileData) & "*.*"
                lDirCount = lDirCount + 1
            End If
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
    
    FileParts = tFS
    
End Function
Public Property Get FileItem(vIndexKey As Variant) As cFiles

  Set FileItem = fpCol(vIndexKey)
  
End Property
Public Property Get FileCount() As Long
    FileCount = lFileCount
End Property
Public Property Get DirItem(vIndexKey As Variant) As cDirs

  Set DirItem = drCol(vIndexKey)
  
End Property
Public Property Get DirCount() As Long
    DirCount = lDirCount
End Property

