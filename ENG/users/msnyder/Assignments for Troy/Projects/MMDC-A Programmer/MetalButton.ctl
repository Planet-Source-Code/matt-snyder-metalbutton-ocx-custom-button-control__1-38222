VERSION 5.00
Begin VB.UserControl MetalButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   MaskColor       =   &H00000000&
   MaskPicture     =   "MetalButton.ctx":0000
   ScaleHeight     =   375
   ScaleWidth      =   1920
   Begin VB.Timer tmrDoStuff 
      Interval        =   10
      Left            =   1200
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   250
      TabIndex        =   0
      Top             =   90
      Width           =   1545
   End
   Begin VB.Image imgButtonClick 
      Height          =   375
      Index           =   2
      Left            =   0
      Picture         =   "MetalButton.ctx":25C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonOver 
      Height          =   375
      Index           =   2
      Left            =   0
      Picture         =   "MetalButton.ctx":4B84
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonNormal 
      Height          =   375
      Index           =   2
      Left            =   0
      Picture         =   "MetalButton.ctx":7146
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonClick 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "MetalButton.ctx":9708
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonOver 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "MetalButton.ctx":BCCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonNormal 
      Height          =   375
      Index           =   1
      Left            =   0
      Picture         =   "MetalButton.ctx":E28C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonNormal 
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "MetalButton.ctx":1084E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonOver 
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "MetalButton.ctx":11910
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
   Begin VB.Image imgButtonClick 
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "MetalButton.ctx":129D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "MetalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'_ ************************************************************** _
 _ *                                                            * _
 _ *      Project:  MetalButton OCX - Custom Button Control     * _
 _ *       Author:  Matt Snyder                                 * _
 _ *      Revised:  08/23/02                                    * _
 _ *                                                            * _
 _ *  Information:  A 3D interactive button                     * _
 _ *                                                            * _
   **************************************************************
Option Explicit
Dim i As Integer
Private propCaption As String
Private propColor As Integer
Private propMaskColor As ColorConstants
Private propFont As StdFont
Dim buttonState As Integer
Dim currentControlState As Integer
Dim oldControlState As Integer
Dim didClick As Boolean
Dim buttonHover As Boolean
Dim CurrenthWnd As Long
Dim Status As Long
Dim scanX As Integer
Dim scanY As Integer
Dim pixelColor As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint& Lib "user32" (ByVal lpPointX As Long, ByVal lpPointY As Long)

Private Const WIDTH_BOUND As Integer = 375
Private Const HEIGHT_BOUND As Integer = 120

Private Const NORMAL_STATE As Integer = 0
Private Const HOVER_STATE As Integer = 1
Private Const PRESSED_STATE As Integer = 2

Private Const LEFT_BUTTON As Byte = 1
Private Const NO_PRESS As Long = 0
Private Const SHORT_PRESS As Long = -32767
Private Const LONG_PRESS As Long = -32768
Private Type POINTAPI
    Xpos As Long
    Ypos As Long
End Type
Dim CursorPos As POINTAPI
   
Public Event Click()
Private Sub UserControl_Initialize()
    Set propFont = New StdFont
    UserControl.MaskColor = RGB(255, 255, 255)
    
    Call ClearButtons
    UserControl.Picture = imgButtonNormal(ButtonColor).Picture
    UserControl.MaskPicture = imgButtonNormal(ButtonColor).Picture
    Caption = "MetalButton"
    Label1.Caption = Caption
End Sub
Private Sub tmrDoStuff_Timer()
    Status = GetCursorPos&(CursorPos)
    CurrenthWnd = WindowFromPoint(CursorPos.Xpos, CursorPos.Ypos)
    Status = GetAsyncKeyState(LEFT_BUTTON)
    
    If CurrenthWnd = UserControl.hWnd And (Status = SHORT_PRESS Or Status = LONG_PRESS) Then
        currentControlState = PRESSED_STATE
        If currentControlState <> oldControlState Then
            Call ClearButtons
            UserControl.Picture = imgButtonClick(ButtonColor).Picture
            UserControl.MaskPicture = imgButtonClick(ButtonColor).Picture
            didClick = True
        End If
        oldControlState = PRESSED_STATE
    ElseIf CurrenthWnd = UserControl.hWnd And didClick = False Then
        currentControlState = HOVER_STATE
        If currentControlState <> oldControlState Then
            Call ClearButtons
            UserControl.Picture = imgButtonOver(ButtonColor).Picture
            UserControl.MaskPicture = imgButtonOver(ButtonColor).Picture
        End If
        oldControlState = HOVER_STATE
    ElseIf CurrenthWnd = UserControl.hWnd And didClick = True And Status = NO_PRESS Then
        didClick = False
        RaiseEvent Click
    Else
        currentControlState = NORMAL_STATE
        If currentControlState <> oldControlState Then
            Call ClearButtons
            UserControl.Picture = imgButtonNormal(ButtonColor).Picture
            UserControl.MaskPicture = imgButtonNormal(ButtonColor).Picture
            didClick = False
        End If
        oldControlState = NORMAL_STATE
    End If
End Sub
Private Sub ClearButtons()
    For i = 0 To 2
        imgButtonNormal(i).Visible = False
        imgButtonOver(i).Visible = False
        imgButtonClick(i).Visible = False
    Next i
End Sub
Private Sub UserControl_Resize()
    For i = 0 To 2
        imgButtonNormal(i).Width = UserControl.Width: imgButtonNormal(i).Height = UserControl.Height
        imgButtonOver(i).Width = UserControl.Width: imgButtonOver(i).Height = UserControl.Height
        imgButtonClick(i).Width = UserControl.Width: imgButtonClick(i).Height = UserControl.Height
    Next i
    Label1.Width = UserControl.Width - WIDTH_BOUND: Label1.Height = UserControl.Height - HEIGHT_BOUND
End Sub
Public Property Get ButtonColor() As Integer
   ButtonColor = propColor
End Property
Public Property Let ButtonColor(ByVal NewColor As Integer)
   If NewColor > 2 Then NewColor = 2
   propColor = NewColor
   PropertyChanged "ButtonColor"
   UserControl.Picture = imgButtonNormal(ButtonColor).Picture
   UserControl.MaskPicture = imgButtonNormal(ButtonColor).Picture
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Properties"
   Caption = propCaption
End Property
Public Property Let Caption(ByVal NewCaption As String)
   propCaption = NewCaption
   PropertyChanged "Caption"
   Label1.Caption = Caption
End Property
Public Property Get Font() As StdFont
   Set Font = propFont
End Property
Public Property Set Font(NewFont As StdFont)
   With propFont
      .Bold = NewFont.Bold
      .Italic = NewFont.Italic
      .Name = NewFont.Name
      .Size = NewFont.Size
   End With
   PropertyChanged "Font"
   
   With Label1.Font
      .Bold = propFont.Bold
      .Italic = propFont.Italic
      .Name = propFont.Name
      .Size = propFont.Size
   End With
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", Caption, "MetalButton"
    PropBag.WriteProperty "ButtonColor", ButtonColor, 0
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "MetalButton")
    ButtonColor = PropBag.ReadProperty("ButtonColor", 0)
End Sub
