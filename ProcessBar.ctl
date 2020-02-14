VERSION 5.00
Begin VB.UserControl ProcessBar 
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ScaleHeight     =   1050
   ScaleWidth      =   3720
   Begin VB.Label lblBase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblTop 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "ProcessBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FigureStyle
    eOnlyValueInt
    eOnlyValueDouble
    ePercentInt
    ePercentDouble
End Enum

Private dblMyPercent As Double
Private lngMyValue As Long
Private isShowMsg As Boolean
Private strMsg As String
Private intMsgStyle As FigureStyle

Private Max As Long, Min As Long

Private Sub UserControl_Initialize()
    Min = 0
    Max = 100
'    lngMyValue = 30
    dblMyPercent = lngMyValue / (Max - Min)
    lblValue.Caption = Int(dblMyPercent * 100) & "/100"
    UserControl.Height = lblBase.Height
    intMsgStyle = ePercentInt
    
    Call setLabelSize
End Sub

Private Sub UserControl_Resize()
    Call setLabelSize
End Sub

Private Sub setLabelSize()
    lblBase.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblTop.Move lblBase.Left, lblBase.Top, dblMyPercent * lblBase.Width, lblBase.Height
    lblValue.Move lblBase.Left, (lblBase.Height - lblValue.Height) / 2 + 20, lblBase.Width
End Sub
'set the value
Public Property Let Value(ByVal Value As Long)
    If Value < Min Then
        lngMyValue = Min
    ElseIf Value > Max Then
        lngMyValue = Max
    Else
        lngMyValue = Value
    End If
    
    dblMyPercent = lngMyValue / (Max - Min)
    lblTop.Width = dblMyPercent * lblBase.Width
    Select Case intMsgStyle
        Case eOnlyValueInt: strMsg = Int(dblMyPercent * 100)
        Case eOnlyValueDouble: strMsg = CStr(dblMyPercent * 100)
        Case ePercentInt: strMsg = Int(dblMyPercent * 100) & "/100"
        Case ePercentDouble: strMsg = CStr(dblMyPercent * 100) & "/100"
    End Select
    lblValue.Caption = strMsg
End Property
'
Public Property Get Value() As Long
    Value = lngMyValue
End Property
'set the percent
Public Property Let Percent(ByVal Percent As Double)
    If Percent < 0 Then
        dblMyPercent = 0
    ElseIf Percent > 1 Then
        dblMyPercent = 1
    Else
        dblMyPercent = Percent
    End If

    lngMyValue = dblMyPercent * (Max - Min)
    lblTop.Width = dblMyPercent * lblBase.Width
    Select Case intMsgStyle
        Case eOnlyValueInt: strMsg = Int(dblMyPercent * 100)
        Case eOnlyValueDouble: strMsg = CStr(dblMyPercent * 100)
        Case ePercentInt: strMsg = Int(dblMyPercent * 100) & "/100"
        Case ePercentDouble: strMsg = CStr(dblMyPercent * 100) & "/100"
    End Select
    lblValue.Caption = strMsg
End Property

Public Property Get Percent() As Double
    Percent = dblMyPercent
End Property
'is show msg
Public Property Let ShowMsg(ByVal ShowMsg As Boolean)
    isShowMsg = ShowMsg
    lblValue.Visible = isShowMsg
End Property

Public Property Get ShowMsg() As Boolean
    ShowMsg = isShowMsg
End Property
'msg style
Public Property Let MsgStyle(ByVal MsgStyle As FigureStyle)
    intMsgStyle = MsgStyle
    Select Case intMsgStyle
        Case eOnlyValueInt: strMsg = Int(dblMyPercent * 100)
        Case eOnlyValueDouble: strMsg = CStr(dblMyPercent * 100)
        Case ePercentInt: strMsg = Int(dblMyPercent * 100) & "/100"
        Case ePercentDouble: strMsg = CStr(dblMyPercent * 100) & "/100"
    End Select
    lblValue.Caption = strMsg
End Property

Public Property Get MsgStyle() As FigureStyle
    MsgStyle = intMsgStyle
End Property
