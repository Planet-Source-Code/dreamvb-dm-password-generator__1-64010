VERSION 5.00
Begin VB.UserControl ButtonFx 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
   Begin VB.PictureBox PicIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00D1D8DB&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ButtonFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private ButtonState As Integer, mUp As Integer

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Sub DrawButton(mState As Integer)
Dim mColors(1) As OLE_COLOR
Dim xPos As Long, yPos As Long

    UserControl.Cls
    
    If mState = 0 Then
        mColors(0) = &HD1D8DB
        mColors(1) = &HD1D8DB
    ElseIf mState = 1 Then
        mColors(0) = &HD2BDB6
        mColors(1) = &H6A240A
    ElseIf mState = 2 Then
        mColors(0) = &HB59285
        mColors(1) = &H6A240A
    End If
    
    PicIcon.BackColor = mColors(0)
    
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), mColors(0), BF
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), mColors(1), B
    
    xPos = (UserControl.ScaleWidth - PicIcon.Width) \ 2
    yPos = (UserControl.ScaleHeight - PicIcon.Height) \ 2
    
    TransparentBlt UserControl.hdc, xPos, yPos, PicIcon.Width, PicIcon.Height, PicIcon.hdc, 0, 0, PicIcon.Width, PicIcon.Height, RGB(255, 0, 255)
    UserControl.Refresh
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button <> vbLeftButton Then Exit Sub
    mUp = 1
    DrawButton 2
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)

    If mUp = 1 Then DrawButton 1: Exit Sub

    If (X < 0) Or (X > UserControl.ScaleWidth) Or (Y < 0) Or (Y > UserControl.ScaleHeight) Then
        ReleaseCapture
        DrawButton 0
    ElseIf GetCapture() <> UserControl.hwnd Then
        DrawButton 1
        SetCapture UserControl.hwnd
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    mUp = 0
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_Resize()
    DrawButton ButtonState
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = PicIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicIcon.Picture = New_Picture
    PropertyChanged "Picture"
    Call DrawButton(ButtonState)
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Show()
    DrawButton ButtonState
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

