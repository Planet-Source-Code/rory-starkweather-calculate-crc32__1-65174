Attribute VB_Name = "mdlDeclarations"
Option Explicit

Public Const OFS_MAXPATHNAME As Long = 128

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

' OpenFile() Flags
Public Const OF_READ = &H0
Public Const OF_WRITE = &H1
Public Const OF_READWRITE = &H2
Public Const OF_SHARE_COMPAT = &H0
Public Const OF_SHARE_EXCLUSIVE = &H10
Public Const OF_SHARE_DENY_WRITE = &H20
Public Const OF_SHARE_DENY_READ = &H30
Public Const OF_SHARE_DENY_NONE = &H40
Public Const OF_PARSE = &H100
Public Const OF_DELETE = &H200
Public Const OF_VERIFY = &H400
Public Const OF_CANCEL = &H800
Public Const OF_CREATE = &H1000
Public Const OF_PROMPT = &H2000
Public Const OF_EXIST = &H4000
Public Const OF_REOPEN = &H8000

Public Declare Function OpenFile _
   Lib "kernel32" _
   (ByVal lpFileName As String, _
    lpReOpenBuff As OFSTRUCT, _
    ByVal wStyle As Long) _
   As Long
   
Public Declare Function ReadFile _
   Lib "kernel32" _
   (ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) _
   As Long
   
Public Declare Function CloseHandle _
   Lib "kernel32" _
   (ByVal hObject As Long) _
   As Long
   
Public Declare Function GetFileSize _
   Lib "kernel32" _
   (ByVal hFile As Long, _
    lpFileSizeHigh As Long) _
   As Long
