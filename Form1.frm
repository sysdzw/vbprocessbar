VERSION 5.00
Object = "{97FA94DA-619B-4FAA-ADBF-9734314E8DFD}#1.0#0"; "ProcessBarPro.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin ProcessBarPro.ProcessBar ProcessBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2400
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Double

Private Sub Command1_Click()
    i = 0
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    i = i + 0.05
    ProcessBar1.Percent = i
    If i > 1 Then Timer1.Enabled = False
End Sub
