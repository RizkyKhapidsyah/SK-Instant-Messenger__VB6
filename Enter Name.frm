VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000C&
   Caption         =   "Enter Name"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "Enter Name.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1200
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H000000FF&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Enter your name here"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtName 
         BackColor       =   &H80000007&
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   120
         MaxLength       =   18
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOk_Click()
Form2.Visible = False
Form1.Show

End Sub
