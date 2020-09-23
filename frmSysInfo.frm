VERSION 5.00
Begin VB.Form frmSysInfo 
   Caption         =   "System Information"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGetSysInfo 
      Caption         =   "&Show System Information"
      Height          =   375
      Left            =   315
      TabIndex        =   1
      Top             =   225
      Width           =   5055
   End
   Begin VB.TextBox txtSysInfo 
      Alignment       =   2  'Center
      Height          =   3210
      Left            =   315
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   5100
   End
End
Attribute VB_Name = "frmSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objSysInfo As New cSystemInfo


Private Sub cmdGetSysInfo_Click()

    Dim strSys As String
    
    With objSysInfo
        strSys = "Your System Information : " & vbCrLf & _
                    "CPU: " & .CPUVersion & vbCrLf & _
                    "IE: " & .IEVersion & vbCrLf & _
                    "MEMORY FREE: " & .MemoryFree & vbCrLf & _
                    "MEMORY TOTAL: " & .MemoryTotal & vbCrLf & _
                    "VIRTUAL FREE: " & .VirtualMemoryFree & vbCrLf & _
                    "VIRTUAL TOTAL: " & .VirtualMemoryTotal & vbCrLf & _
                    "WINDOWS VER: " & .WinName & "  " & .WinVersion & vbCrLf
                    
    End With
    
    Me.txtSysInfo.Text = strSys

End Sub
