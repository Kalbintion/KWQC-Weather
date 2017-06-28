VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KWQC Weather - By: Anthoni Wiese"
   ClientHeight    =   5205
   ClientLeft      =   8835
   ClientTop       =   6420
   ClientWidth     =   7575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   Begin VB.Timer tmrGetPic 
      Interval        =   1000
      Left            =   2640
      Top             =   3480
   End
   Begin VB.Image imgCur 
      Height          =   5100
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
   Begin VB.Menu mnuPicture 
      Caption         =   "Picture"
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Image"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPic 
         Caption         =   "7-Day"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Radar"
         Index           =   1
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Radar (Loop)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Current"
         Index           =   3
      End
      Begin VB.Menu mnuPic 
         Caption         =   "Current Area"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwRserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private URLs(0 To 4) As String
Private IMGs(0 To 4) As String
Private refreshCounter As Long

Private Sub Form_Load()
    URLs(0) = "http://ftpcontent.worldnow.com/kwqc/WEATHER_7day.jpg"
    IMGs(0) = "cur7Day.kdk"
    
    URLs(1) = "http://ftpcontent.worldnow.com/kwqc/WEATHER_radar.jpg"
    IMGs(1) = "curRadarStill.kdk"
    
    URLs(2) = "http://ftpcontent.worldnow.com/kwqc/WEATHER_radar_loop.gif"
    IMGs(2) = "curRadarLoop.kdk"
    
    URLs(3) = "http://ftpcontent.worldnow.com/kwqc/WEATHER_currents.jpg"
    IMGs(3) = "curNow.kdk"
    
    URLs(4) = "http://ftpcontent.worldnow.com/kwqc/WEATHER_Temps.jpg"
    IMGs(4) = "curNowArea.kdk"
    
    GetImageUpdate
    
    GetSelection
End Sub

Private Sub imgCur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbKeyRButton Then
        PopupMenu mnuPicture
    End If
End Sub

Private Sub mnuOptions_Click()
    frmOpts.Show
End Sub

Private Sub mnuPic_Click(Index As Integer)
    For i = 0 To mnuPic.UBound
        mnuPic(i).Checked = False
    Next
    mnuPic(Index).Checked = True
    
    GetSelection
End Sub

Private Sub mnuRefresh_Click()
    GetImageUpdate
End Sub

Private Sub tmrGetPic_Timer()
    refreshCounter = refreshCounter + 1
    If refreshCounter = frmOpts.hsTime.Value Then
        GetImageUpdate
    End If
End Sub

Private Sub GetSelection()
    On Error GoTo loopTilFilesDone


    Dim selCode As Long
    For i = 0 To mnuPic.UBound
        If mnuPic(i).Checked = True Then
            selCode = i
            Exit For
        End If
    Next
    
    imgCur.Picture = Nothing
    
loopTilFilesDone:
    Select Case selCode
        Case 0:
            ' 7-Day
            imgCur.Picture = LoadPicture(App.Path & "\cur7Day.kdk")
        Case 1:
            ' Radar Still
            imgCur.Picture = LoadPicture(App.Path & "\curRadarStill.kdk")
        Case 2:
            ' Radar Loop
            imgCur.Picture = LoadPicture(App.Path & "\curRadarLoop.kdk")
        Case 3:
            ' Current
            imgCur.Picture = LoadPicture(App.Path & "\curNow.kdk")
        Case 4:
            ' Current Area
            imgCur.Picture = LoadPicture(App.Path & "\curNowArea.kdk")
    End Select
End Sub

Private Sub GetImageUpdate()
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

Private Function FileExists(fileName As String) As Boolean
    On Error GoTo errorhandler
    FileExists = (GetAttr(fileName) And vbDirectory) = 0
errorhandler:
    Exit Function
End Function

Private Sub WriteError(errCode As Long, class As Long)
    Dim fNum As Long, fPath As String
    fNum = FreeFile
    fPath = App.Path & "\err.kdk"
    Open fPath For Append As fNum
    Print #fNum, Now & vbTab & errCode & vbTab & class
    Close #fNum
End Sub


