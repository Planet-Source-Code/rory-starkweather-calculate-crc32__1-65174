Attribute VB_Name = "mdlProcs"
Option Explicit

'***********************************************************************************
'***  Procedure: AddSlash
'***
'***  Purpose: Adds a \ to a path if it doesn't have one.
'***
'***  Inputs: strPath: The path.
'***
'***  Outputs: String: Path with trailing slash.
'***
'***  Last Modification: 05/01/2006
'***********************************************************************************
Public Function AddSlash(strPath As String) As String
   If Right$(strPath, 1) <> "\" Then
      strPath = strPath & "\"
   End If
   AddSlash = strPath
End Function

Public Function GetFileQuick(strFilePath As String) As String

  Dim arrFileMain() As Byte
  Dim lngFileSize As Long
  Dim lngRetVal As Long
  Dim lngFileHandle As Long
  Dim ofData As OFSTRUCT

   'Open the two files
   lngFileHandle = OpenFile(strFilePath, ofData, OF_READ)
   
   'Get the file size
   lngFileSize = GetFileSize(lngFileHandle, 0)
   
   'Create an array of bytes
   ReDim arrFileMain(lngFileSize) As Byte
   
   'Read from the file
   lngRetVal = ReadFile(lngFileHandle, _
                          arrFileMain(0), _
                          UBound(arrFileMain), _
                          lngRetVal, _
                          ByVal 0&)
   
   'Close the file
   lngRetVal = CloseHandle(lngFileHandle)
   
   '*** Remember that the array is zero based but the
   '*** file length is one based.
   ReDim Preserve arrFileMain(UBound(arrFileMain) - 1)
   
   GetFileQuick = StrConv(arrFileMain(), vbUnicode)

End Function

'***********************************************************************************
'***  Procedure: FileExists
'***
'***  Purpose: Tests file specs for existence
'***
'***  Inputs: strFileSpec: File spec in question.
'***
'***  Outputs: Boolean exist/doesn't exist.
'***
'***  Last Modification: 04/30/2006
'***********************************************************************************
Public Function FileExists(strFileSpec As String) As Boolean

   Dim blnReturnValue As Boolean
   
   If Dir(strFileSpec) = vbNullString Then
      blnReturnValue = False
   Else
      blnReturnValue = True
   End If
   
   FileExists = blnReturnValue
   
End Function
