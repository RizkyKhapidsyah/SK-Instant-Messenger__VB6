VERSION 5.00
Object = "{1D8A3351-C678-11D1-AA6F-000000000000}#1.0#0"; "DARTSOCK.DLL"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Instant Messenger"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   Icon            =   "Instant Messenger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Incoming 
      BackColor       =   &H8000000C&
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtIn 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000009&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin DartSockCtl.Udp Udp1 
         Left            =   3000
         OleObjectBlob   =   "Instant Messenger.frx":1272
         Top             =   2520
      End
   End
   Begin VB.Frame Outgoing 
      BackColor       =   &H8000000C&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
      Begin VB.TextBox txtOut 
         BackColor       =   &H80000007&
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const pad = 75
Const ChatPort = 51345





Private Sub Form_Activate()
    txtOut.SetFocus
End Sub

Private Sub Form_Load()
Form2.Show
Form1.Visible = False
    Udp1.Open ChatPort
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.Height < 3500 Then Me.Height = 3500
    If Me.Width < 5900 Then Me.Width = 5900
    
    With Incoming
        .Top = pad
        .Left = pad
        .Width = ScaleWidth - (2 * pad)
        .Height = ScaleHeight - Outgoing.Height - (3 * pad)
    End With
        
    With Outgoing
        .Top = Incoming.Top + Incoming.Height + pad
        .Left = Incoming.Left
        .Width = Incoming.Width
    End With
    
    With txtIn
        .Top = pad * 3
        .Left = pad
        .Width = Incoming.Width - (2 * pad)
        .Height = Incoming.Height - (4 * pad)
    End With
    
    With txtOut
        .Top = pad * 3
        .Left = pad
        .Width = Outgoing.Width - (2 * pad)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub






Private Sub txtIn_KeyPress(KeyAscii As Integer)
   
    KeyAscii = 0
End Sub

Private Sub txtOut_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Udp1.Send "[" + Form2.txtName.Text + "]: " + txtOut.Text, "192.168.1.1", ChatPort
        txtOut.Text = ""
    End If
End Sub

Private Sub Udp1_Error(ByVal Number As DartSockCtl.ErrorConstants, ByVal Description As String)
    If Number = 10048 Then
        MsgBox "UDP Port " + CStr(ChatPort) + " is already in Use!" _
        + vbCrLf + "Please close the application that is " _
        + "using Port " + CStr(ChatPort) + " and try again." _
        , , "PowerTCP Error"
        Unload Me
    Else
        MsgBox "Error: " + Description, , "PowerTCP Error"
    End If
End Sub

Private Sub Udp1_Receive()
    Dim s As String
    Udp1.Receive s
    txtIn.SelText = s
    On Error Resume Next
    txtOut.SetFocus
End Sub


