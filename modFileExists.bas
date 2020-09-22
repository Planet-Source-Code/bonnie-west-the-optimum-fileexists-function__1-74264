Attribute VB_Name = "modFileExists"
'=========================================================================
'                     The Optimum FileExists Function
'
'The ideal implementation of FileExists should be simple, efficient,
'supports wildcards and above all else, work flawlessly in all scenarios.
'In the refined to near perfection version 11.0 below, all of those are
'met, except one. For that single shortcoming, v7.0 fills the role
'adequately.
'
'Bonus: A few related routines are included as well.
'=========================================================================

Option Explicit

Private Const DRIVE_NO_ROOT_DIR       As Long = 1
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Const MAX_PATH                As Long = 260

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As Long _
) As Long
Private Declare Function FindFirstFileW Lib "kernel32" ( _
    ByVal lpFileName As Long, _
    ByRef lpFindFileData As WIN32_FIND_DATA _
) As Long


Private Declare Function GetDriveTypeW Lib "kernel32" ( _
    ByVal lpRootPathName As Long _
) As Long


Private Declare Function GetFileAttributesW Lib "kernel32" ( _
    ByVal lpFileName As Long _
) As Long


Private Declare Function PathFileExistsW Lib "shlwapi" ( _
    ByVal pszPath As Long _
) As Long
Private Declare Function PathIsDirectoryW Lib "shlwapi" ( _
    ByVal pszPath As Long _
) As Long

'=========================================================================

Public Function FileExists(ByRef sFileName As String) As Boolean
'·······························   v1.0   ································
'
'Naive beginner's initial attempt.
'
'If Dir$(sFileName, vbArchive) = "" And _
'   Dir$(sFileName, vbHidden) = "" And _
'   Dir$(sFileName, vbReadOnly) = "" And _
'   Dir$(sFileName, vbSystem) = "" Then
'    FileExists = False
'Else
'    FileExists = True
'End If
'
'·······························   v2.0   ································
'
'One-liner form of the above. Unwittingly made worse by use of IIf.
'
'FileExists = IIf(Dir$(sFileName, vbArchive) = "" And _
'                 Dir$(sFileName, vbHidden) = "" And _
'                 Dir$(sFileName, vbReadOnly) = "" And _
'                 Dir$(sFileName, vbSystem) = "", False, True)
'
'·······························   v3.0   ································
'
'Code inspired by Kevin Wilson (www.thevbzone.com) & Francesco Balena.
'http://www.thevbzone.com/modCommon.bas
'
'FileExists = Dir$(sFileName, vbArchive Or vbHidden Or _
'                            vbReadOnly Or vbSystem) <> ""
'
'·······························   v4.0   ································
'
'Exits early if sFileName is empty, returning the default value False.
'
'On Error Resume Next
'If LenB(sFileName) Then _
'    FileExists = Dir$(sFileName, vbArchive Or vbHidden Or _
'                                vbReadOnly Or vbSystem) <> vbNullString
'
'·······························   v5.0   ································
'
'Rejects Directories/Folders, returning the default value False.
'
'On Error Resume Next
'If LenB(sFileName) Then If Right$(sFileName, 1) <> "\" Then _
'    FileExists = Dir$(sFileName, vbArchive Or vbHidden Or _
'                                vbReadOnly Or vbSystem) <> vbNullString
'
'·······························   v6.0   ································
'
'Doesn't accept wildcards. Opening a locked file fails.
'"Close FreeFile - 1" may not always work as expected.
'
'On Error Resume Next
'Open sFileName For Input As FreeFile
'    FileExists = (Err = 0)
'Close FreeFile - 1
'
'·······························   v7.0   ································
'
'Wide version of FindFirstFile API allows Unicode filenames and makes
'passing the string faster thus contributing to the overall efficiency
'of this code. Supports wildcards.
'
'Dim WFD As WIN32_FIND_DATA
'If LenB(sFileName) Then _
'    FileExists = FindClose(FindFirstFileW(StrPtr(sFileName), WFD)) <> 0
'
'·······························   v8.0   ································
'
'GetAttr throws an error with empty strings, wildcards, hiberfil.sys,
'pagefile.sys, NUL, CON, COM1, etc. thus causing False to be returned.
'Directories/Folders are excluded by the test.
'
'On Error Resume Next
'FileExists = (GetAttr(sFileName) And vbDirectory) <> vbDirectory
'
'·······························   v9.0   ································
'
'Does not recognize wildcards, hiberfil.sys & pagefile.sys, thus returns
'False. Ignores Directories/Folders if assisted by PathIsDirectory
'or similar.
'
'If PathIsDirectoryW(StrPtr(sFileName)) = False Then _
'    FileExists = PathFileExistsW(StrPtr(sFileName))
'
'······························   v10.0   ································
'
'The Scripting version is much slower than any of the others,
'even if it is referenced. Wildcards not supported.
'
'Dim FSO As Object  'Or FSO As New FileSystemObject
'Set FSO = CreateObject("Scripting.FileSystemObject")
'FileExists = FSO.FileExists(sFileName)
'Set FSO = Nothing
'
'······························   v11.0   ································
'
'Wildcards unsupported but this is the fastest file existence test yet.
'Superstition: Why is GetFileAttributes the way old-timers
'test file existence? (by Raymond Chen)
'http://blogs.msdn.com/b/oldnewthing/archive/2007/10/23/5612082.aspx
'Check if a file exists (by Wolfgang Enzinger)
'http://www.enzinger.net/en/Filetest.html
'
Select Case (GetFileAttributesW(StrPtr(sFileName)) And vbDirectory) = 0
    Case True: FileExists = True
    Case Else: FileExists = (Err.LastDllError = ERROR_SHARING_VIOLATION)
End Select
'
'······························   v12.0   ································
'
'This one-liner form of the above does the LastDllError check everytime.
'
'FileExists = ((GetFileAttributesW(StrPtr(sFileName)) And vbDirectory) _
'           = 0) Or (Err.LastDllError = ERROR_SHARING_VIOLATION)
End Function

'=========================================================================

Public Function DirExists(ByRef sPath As String) As Boolean
    DirExists = Abs(GetFileAttributesW(StrPtr(sPath))) And vbDirectory
End Function

Public Function DriveExists(ByRef sDrive As String) As Boolean
    DriveExists = GetDriveTypeW(StrPtr(sDrive)) <> DRIVE_NO_ROOT_DIR
End Function

Public Function GetVolumeLabel(ByRef sDrive As String) As String
    GetVolumeLabel = Dir$(sDrive, vbVolume)
End Function
