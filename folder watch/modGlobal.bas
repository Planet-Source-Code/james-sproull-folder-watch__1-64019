Attribute VB_Name = "modGlobal"
Public strReportFile As String, intSlashPos As Integer
Public intCnt As Integer
Public FSO As New FileSystemObject
Public strFolder As Folder, strSubFolder As Folder, strFile As File
Public strScanDir As String
Public strpath As String, strDsc As Boolean
Public intTotFiles As Integer, intTotDirs As Integer
Public intCurDepth As Integer, strInitDir As String

Public Function ExtractExt(strSource As String) As Boolean
Dim strExt As String, intPos As Integer
    intPos = InStrRev(strSource, ".")
    strExt = Mid$(strSource, intPos + 1)
    If strExt = "dsc" Then
        ExtractExt = True
    Else
        ExtractExt = False
    End If
End Function

