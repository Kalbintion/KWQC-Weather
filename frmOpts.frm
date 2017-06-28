VERSION 5.00
Begin VB.Form frmOpts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KWQC Weather - Options"
   ClientHeight    =   1200
   ClientLeft      =   1875
   ClientTop       =   2715
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3840
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtRefreshSec 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtRefreshMin 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.HScrollBar hsTime 
      Height          =   255
      LargeChange     =   30
      Left            =   120
      Max             =   600
      Min             =   60
      TabIndex        =   0
      Top             =   480
      Value           =   60
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Refresh Time"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    Dim fNum As Long, fPath As String
    fPath = App.Path & "\settings.kdk"
    fNum = FreeFile
    Open fPath For Output As fNum
    Print #fNum, hsTime.Value
    Close #fNum
    
    Me.Visible = False
End Sub

Private Sub hsTime_Change()
    txtRefreshSec.Text = hsTime.Value Mod 60
    txtRefreshMin.Text = hsTime.Value / 60 Mod 60
End Sub

Private Sub txtRefreshMin_Change()
    On Error Resume Next
    For i = 0 To 255
        If i <> Asc("0") And i <> Asc("1") And i <> Asc("2") And i <> Asc("3") And i <> Asc("4") And i <> Asc("5") And i <> Asc("6") And i <> Asc("7") And i <> Asc("8") And i <> Asc("9") Then
            txtRefreshMin.Text = Replace$(txtRefreshMin.Text, Chr(i), "")
        End If
    Next
    If Len(txtRefreshMin.Text) = 0 Then txtRefreshMin.Text = 1
    If Val(txtRefreshMin.Text) < 0 Then txtRefreshMin.Text = 0
    If Val(txtRefreshMin.Text) > 10 Then txtRefreshMin.Text = 10
    hsTime.Value = Val(txtRefreshMin.Text) * 60 + Val(txtRefreshSec.Text)
End Sub

Private Sub txtRefreshSec_Change()
    On Error Resume Next
    For i = 0 To 255
        If i <> Asc("0") And i <> Asc("1") And i <> Asc("2") And i <> Asc("3") And i <> Asc("4") And i <> Asc("5") And i <> Asc("6") And i <> Asc("7") And i <> Asc("8") And i <> Asc("9") Then
            txtRefreshSec.Text = Replace$(txtRefreshSec.Text, Chr(i), "")
        End If
    Next
    If Len(txtRefreshSec.Text) = 0 Then txtRefreshMin.Text = 1
    If Val(txtRefreshSec.Text) < 0 Then txtRefreshSec.Text = 0
    If Val(txtRefreshSec.Text) > 59 Then txtRefreshSec.Text = 59
    hsTime.Value = Val(txtRefreshMin.Text) * 60 + Val(txtRefreshSec.Text)
End Sub
