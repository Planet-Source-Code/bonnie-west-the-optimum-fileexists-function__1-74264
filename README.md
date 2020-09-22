<div align="center">

## The Optimum FileExists Function


</div>

### Description

<P>

The ideal implementation of FileExists should be simple, efficient,

supports wildcards and above all else, work flawlessly in all scenarios.

In the refined to near perfection version 11.0 below, all of those are

met, except one. For that single shortcoming, v7.0 fills the role

adequately.

</P><P>

Bonus: A few related routines are included as well.

</P>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2012-03-03 12:30:02
**By**             |[Bonnie West](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bonnie-west.md)
**Level**          |Beginner
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[The\_Optimu222104332012\.zip](https://github.com/Planet-Source-Code/bonnie-west-the-optimum-fileexists-function__1-74264/archive/master.zip)





### Source Code

<PRE><FONT color="#000080">Attribute</FONT> VB_Name = <FONT color="#800000">"modFileExists"</FONT><FONT color="#008000">
'=========================================================================
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The Optimum FileExists Function
'
'The ideal implementation of FileExists should be simple, efficient,
'supports wildcards and above all else, work flawlessly in all scenarios.
'In the refined to near perfection version 11.0 below, all of those are
'met, except one. For that single shortcoming, v7.0 fills the role
'adequately.
'
'Bonus: A few related routines are included as well.
'=========================================================================</FONT>
<FONT color="#000080">
Option Explicit
</FONT>
<FONT color="#000080">Private Const</FONT> DRIVE_NO_ROOT_DIR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As Long</FONT> = <FONT color="#800080">1</FONT>
<FONT color="#000080">Private Const</FONT> ERROR_SHARING_VIOLATION <FONT color="#000080">As Long</FONT> = <FONT color="#800080">32</FONT>
<FONT color="#000080">Private Const</FONT> MAX_PATH&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As Long</FONT> = <FONT color="#800080">260</FONT>
<FONT color="#000080">
Private Type</FONT> FILETIME
&nbsp;&nbsp;&nbsp;&nbsp;dwLowDateTime&nbsp;&nbsp;<FONT color="#000080">As Long</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;dwHighDateTime <FONT color="#000080">As Long
End Type
</FONT><FONT color="#000080">
Private Type</FONT> WIN32_FIND_DATA
&nbsp;&nbsp;&nbsp;&nbsp;dwFileAttributes <FONT color="#000080">As Long</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;ftCreationTime&nbsp;&nbsp;&nbsp;<FONT color="#000080">As</FONT> FILETIME
&nbsp;&nbsp;&nbsp;&nbsp;ftLastAccessTime <FONT color="#000080">As</FONT> FILETIME
&nbsp;&nbsp;&nbsp;&nbsp;ftLastWriteTime&nbsp;&nbsp;<FONT color="#000080">As</FONT> FILETIME
&nbsp;&nbsp;&nbsp;&nbsp;nFileSizeHigh&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As Long</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;nFileSizeLow&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As Long</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;dwReserved0&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As Long</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;dwReserved1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As Long</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;cFileName&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As String</FONT> * MAX_PATH
&nbsp;&nbsp;&nbsp;&nbsp;cAlternate&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">As String</FONT> * <FONT color="#800080">14</FONT>
<FONT color="#000080">End Type
</FONT><FONT color="#000080">
Private Declare Function</FONT> FindClose <FONT color="#000080">Lib</FONT> <FONT color="#800000">"kernel32"</FONT> ( _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByVal</FONT> hFindFile <FONT color="#000080">As Long</FONT> _
) <FONT color="#000080">As Long
Private Declare Function</FONT> FindFirstFileW <FONT color="#000080">Lib</FONT> <FONT color="#800000">"kernel32"</FONT> ( _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByVal</FONT> lpFileName <FONT color="#000080">As Long</FONT>, _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByRef</FONT> lpFindFileData <FONT color="#000080">As</FONT> WIN32_FIND_DATA _
) <FONT color="#000080">As Long
</FONT>
<FONT color="#000080">
Private Declare Function</FONT> GetDriveTypeW <FONT color="#000080">Lib</FONT> <FONT color="#800000">"kernel32"</FONT> ( _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByVal</FONT> lpRootPathName <FONT color="#000080">As Long</FONT> _
) <FONT color="#000080">As Long
</FONT>
<FONT color="#000080">
Private Declare Function</FONT> GetFileAttributesW <FONT color="#000080">Lib</FONT> <FONT color="#800000">"kernel32"</FONT> ( _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByVal</FONT> lpFileName <FONT color="#000080">As Long</FONT> _
) <FONT color="#000080">As Long
</FONT>
<FONT color="#000080">
Private Declare Function</FONT> PathFileExistsW <FONT color="#000080">Lib</FONT> <FONT color="#800000">"shlwapi"</FONT> ( _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByVal</FONT> pszPath <FONT color="#000080">As Long</FONT> _
) <FONT color="#000080">As Long
Private Declare Function</FONT> PathIsDirectoryW <FONT color="#000080">Lib</FONT> <FONT color="#800000">"shlwapi"</FONT> ( _
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">ByVal</FONT> pszPath <FONT color="#000080">As Long</FONT> _
) <FONT color="#000080">As Long</FONT>
<FONT color="#008000">
'=========================================================================
</FONT><FONT color="#000080">
Public Function</FONT> FileExists(<FONT color="#000080">ByRef</FONT> sFileName <FONT color="#000080">As String</FONT>) <FONT color="#000080">As Boolean</FONT><FONT color="#008000">
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v1.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Naive beginner's initial attempt.
'
'If Dir$(sFileName, vbArchive) = "" And </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;Dir$(sFileName, vbHidden) = "" And </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;Dir$(sFileName, vbReadOnly) = "" And </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;Dir$(sFileName, vbSystem) = "" Then
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = False
'Else
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = True
'End If
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v2.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'One-liner form of the above. Unwittingly made worse by use of IIf.
'
'FileExists = IIf(Dir$(sFileName, vbArchive) = "" And </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dir$(sFileName, vbHidden) = "" And </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dir$(sFileName, vbReadOnly) = "" And </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dir$(sFileName, vbSystem) = "", False, True)
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v3.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Code inspired by Kevin Wilson (<A href="http://www.thevbzone.com/modCommon.bas">www.thevbzone.com</A>) & Francesco Balena.
'
'FileExists = Dir$(sFileName, vbArchive Or vbHidden Or </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;vbReadOnly Or vbSystem) <> ""
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v4.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Exits early if sFileName is empty, returning the default value False.
'
'On Error Resume Next
'If LenB(sFileName) Then </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = Dir$(sFileName, vbArchive Or vbHidden Or </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;vbReadOnly Or vbSystem) <> vbNullString
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v5.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Rejects Directories/Folders, returning the default value False.
'
'On Error Resume Next
'If LenB(sFileName) Then If Right$(sFileName, 1) <> "\" Then </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = Dir$(sFileName, vbArchive Or vbHidden Or </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;vbReadOnly Or vbSystem) <> vbNullString
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v6.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Doesn't accept wildcards. Opening a locked file fails.
'"Close FreeFile - 1" may not always work as expected.
'
'On Error Resume Next
'Open sFileName For Input As FreeFile
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = (Err = 0)
'Close FreeFile - 1
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v7.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Wide version of FindFirstFile API allows Unicode filenames and makes
'passing the string faster thus contributing to the overall efficiency
'of this code. Supports wildcards.
'
'Dim WFD As WIN32_FIND_DATA
'If LenB(sFileName) Then </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = FindClose(FindFirstFileW(StrPtr(sFileName), WFD)) <> 0
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v8.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'GetAttr throws an error with empty strings, wildcards, hiberfil.sys,
'pagefile.sys, NUL, CON, COM1, etc. thus causing False to be returned.
'Directories/Folders are excluded by the test.
'
'On Error Resume Next
'FileExists = (GetAttr(sFileName) And vbDirectory) <> vbDirectory
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v9.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Does not recognize wildcards, hiberfil.sys & pagefile.sys, thus returns
'False. Ignores Directories/Folders if assisted by PathIsDirectory
'or similar.
'
'If PathIsDirectoryW(StrPtr(sFileName)) = False Then </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;FileExists = PathFileExistsW(StrPtr(sFileName))
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v10.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'The Scripting version is much slower than any of the others,
'even if it is referenced. Wildcards not supported.
'
'Dim FSO As Object 'Or FSO As New FileSystemObject
'Set FSO = CreateObject("Scripting.FileSystemObject")
'FileExists = FSO.FileExists(sFileName)
'Set FSO = Nothing
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v11.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'Wildcards unsupported but this is the fastest file existence test yet.
'<A href="http://blogs.msdn.com/b/oldnewthing/archive/2007/10/23/5612082.aspx">Superstition: Why is GetFileAttributes the way old-timers</A>
'<A href="http://blogs.msdn.com/b/oldnewthing/archive/2007/10/23/5612082.aspx">test file existence?</A> (by Raymond Chen)
'<A href="http://www.enzinger.net/en/Filetest.html">Check if a file exists</A> (by Wolfgang Enzinger)
'</FONT>
<FONT color="#000080">Select Case</FONT> (GetFileAttributesW(StrPtr(sFileName)) <FONT color="#000080">And</FONT> vbDirectory) = <FONT color="#800080">0</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">Case True</FONT>: FileExists = <FONT color="#000080">True</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#000080">Case Else</FONT>: FileExists = (Err.LastDllError = ERROR_SHARING_VIOLATION)
<FONT color="#000080">End Select</FONT><FONT color="#008000">
'
'&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&nbsp;&nbsp;&nbsp;v12.0&nbsp;&nbsp;&nbsp;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;&middot;
'
'This one-liner form of the above does the LastDllError check everytime.
'
'FileExists = ((GetFileAttributesW(StrPtr(sFileName)) And vbDirectory) </FONT>_<FONT color="#008000">
'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;= 0) Or (Err.LastDllError = ERROR_SHARING_VIOLATION)</FONT><FONT color="#000080">
End Function
</FONT><FONT color="#008000">
'=========================================================================
</FONT><FONT color="#000080">
Public Function</FONT> DirExists(<FONT color="#000080">ByRef</FONT> sPath <FONT color="#000080">As String</FONT>) <FONT color="#000080">As Boolean</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;DirExists = Abs(GetFileAttributesW(StrPtr(sPath))) <FONT color="#000080">And</FONT> vbDirectory
<FONT color="#000080">End Function
</FONT><FONT color="#000080">
Public Function</FONT> DriveExists(<FONT color="#000080">ByRef</FONT> sDrive <FONT color="#000080">As String</FONT>) <FONT color="#000080">As Boolean</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;DriveExists = GetDriveTypeW(StrPtr(sDrive)) <> DRIVE_NO_ROOT_DIR
<FONT color="#000080">End Function
</FONT><FONT color="#000080">
Public Function</FONT> GetVolumeLabel(<FONT color="#000080">ByRef</FONT> sDrive <FONT color="#000080">As String</FONT>) <FONT color="#000080">As String</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;GetVolumeLabel = Dir$(sDrive, vbVolume)
<FONT color="#000080">End Function
</FONT></PRE>

