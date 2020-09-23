VERSION 5.00
Begin VB.UserControl SpinFx 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   Begin VB.PictureBox PicButton 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   900
      ScaleHeight     =   135
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   135
      Width           =   240
   End
   Begin VB.PictureBox PicButton 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   900
      ScaleHeight     =   135
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   1695
      Picture         =   "SpinFx.ctx":0000
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1245
      Top             =   30
   End
   Begin VB.TextBox txtVal 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   885
   End
End
Attribute VB_Name = "SpinFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim op As Integer, nVal As Integer
Private m_Max As Integer
Event Change()


Sub PaintButton(Index As Integer, Pos As Integer)
    BitBlt PicButton(Index).hdc, 0, 0, 16, 8, PicSrc.hdc, 0, Pos, vbSrcCopy
    PicButton(Index).Refresh
End Sub

Sub PaintButtons(Index As Integer)
    BitBlt PicButton(0).hdc, 0, 0, 16, 8, PicSrc.hdc, 0, 0, vbSrcCopy
    BitBlt PicButton(1).hdc, 0, 0, 16, 8, PicSrc.hdc, 0, 9, vbSrcCopy
    PicButton(Index).Refresh: PicButton(1).Refresh
End Sub

Sub UpdateSpinValue()
    If op = 0 Then
        Timer1.Enabled = False
        Exit Sub
    Else
        Timer1.Enabled = True
    End If
End Sub

Private Sub PicButton_Click(Index As Integer)
    Call UpdateSpinValue
End Sub

Private Sub PicButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> vbLeftButton Then Exit Sub
    If Index = 0 Then PaintButton 0, 18
    If Index = 1 Then PaintButton Index, 27
    '
    op = Index + 1
    UpdateSpinValue
    DoEvents
End Sub

Private Sub PicButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> vbLeftButton Then Exit Sub
    If Index = 0 Then PaintButton 0, 0
    If Index = 1 Then PaintButton Index, 9
    op = 0
    UpdateSpinValue
    
End Sub

Private Sub Timer1_Timer()
    
    nVal = Val(txtVal.Text)
        
    If op = 1 Then
        If nVal = m_Max Then Exit Sub
        nVal = nVal + 1
    ElseIf op = 2 Then
        If nVal <= 0 Then nVal = m_Max
        nVal = nVal - 1
    End If
        
    txtVal.Text = nVal
    
End Sub

Private Sub UserControl_Initialize()
    m_Max = 100
    nVal = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Max", m_Max, 100
    PropBag.WriteProperty "Value", nVal, 0
    Call PropBag.WriteProperty("Text", txtVal.Text, "0")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Max = PropBag.ReadProperty("Max", 100)
   nVal = PropBag.ReadProperty("Value", 0)
    txtVal.Text = PropBag.ReadProperty("Text", "0")
End Sub

Private Sub UserControl_Resize()
    PaintButtons 0
    PicButton(0).Left = (UserControl.ScaleWidth) - PicButton(0).Width
    PicButton(1).Left = (UserControl.ScaleWidth) - PicButton(1).Width
    txtVal.Width = PicButton(0).Left
    UserControl.Height = txtVal.Height * Screen.TwipsPerPixelY + 29
End Sub

Private Sub UserControl_Show()
    PaintButtons 0
    txtVal.Text = nVal
End Sub

Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal vNewMax As Integer)
    m_Max = vNewMax
End Property

Public Property Get Value() As Integer
    Value = nVal
End Property

Public Property Let Value(ByVal vNewValue As Integer)
    nVal = vNewValue
End Property

Private Sub txtVal_Change()
    RaiseEvent Change
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtVal.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtVal.Text() = New_Text
    PropertyChanged "Text"
End Property

