Attribute VB_Name = "Module1"
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwRserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Sub GetImageUpdateEx()
    
End Sub

Public Sub GetImageUpdate()
    Dim errCode As Long
    
    For i = 0 To UBound(IMGs)
        If i = 2 Then GoTo skipFile
        
        If FileExists(App.Path & "\" & IMGs(i)) Then
            Kill App.Path & "\" & IMGs(i)
        End If
        Call DeleteUrlCacheEntry(URLs(i))
        errCode = URLDownloadToFile(0, URLs(i), App.Path & "\" & IMGs(i), 0, 0)
        If errCode <> 0 Then
            Call WriteError(errCode, i + 1)
        End If
skipFile:
        errCode = 0
    Next
    
    GetSelection
End Sub
