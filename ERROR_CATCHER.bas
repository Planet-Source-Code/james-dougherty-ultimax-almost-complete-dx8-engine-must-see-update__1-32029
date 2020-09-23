Attribute VB_Name = "ERROR_CATCHER"
Option Explicit

'HANDLES ERRORS...WRITES REPORT TO APP.PATH IF AN ERROR OCCOURS

Private FSys As New FileSystemObject
Private OutStream As TextStream
Public ErrStr As String
Public PossibleCause As String
Public D3DD As Direct3DDevice8

Public Sub ErrorToFile(OutputFileName As String, StringERROR As String, Optional PossibleCause As String = "")
On Local Error Resume Next
Dim tempStr1 As String
Dim tempStr2 As String

Set OutStream = FSys.CreateTextFile(OutputFileName & ".txt", True, False)

tempStr1 = "There Was An ERROR At - "
tempStr2 = "Possible Cause - "

OutStream.WriteLine tempStr1 & StringERROR
OutStream.WriteLine ""
OutStream.WriteLine "Error - " & Err.Description
OutStream.WriteLine ""
OutStream.WriteLine tempStr2
OutStream.WriteLine PossibleCause

Set OutStream = Nothing
End Sub
