VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "eRay Studios Password Generator"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicInv2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1560
      Picture         =   "frmmain.frx":0E42
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   43
      Top             =   7380
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicInv1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1530
      Picture         =   "frmmain.frx":1184
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   42
      Top             =   7035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "Uppercase [A-Z]"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   4560
      Value           =   -1  'True
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "Lowecase [a-z]"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   4815
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "Numeric [0-9]"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   5100
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "All above [A-Z] [a-z] [0-9]"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   36
      Top             =   5370
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "Extanted charset [127-255]"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   5625
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "Special chars [#+=-*+/@^]"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   5865
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "All of above"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Width           =   2235
   End
   Begin VB.OptionButton OptGen 
      Caption         =   "Binary Random [0-1]"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   32
      Top             =   6375
      Width           =   2235
   End
   Begin VB.PictureBox PicA 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   2790
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   299
      TabIndex        =   30
      Top             =   1230
      Width           =   4485
      Begin VB.Label lblGen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.ComboBox cboSort 
      Height          =   315
      ItemData        =   "frmmain.frx":14C6
      Left            =   135
      List            =   "frmmain.frx":14C8
      TabIndex        =   29
      Top             =   3570
      Width           =   1665
   End
   Begin VB.CheckBox ChkDup 
      Caption         =   "Generate No Duplications"
      Height          =   210
      Left            =   135
      TabIndex        =   27
      Top             =   3030
      Value           =   1  'Checked
      Width           =   2220
   End
   Begin VB.PictureBox rev2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   810
      Picture         =   "frmmain.frx":14CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      Top             =   7365
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Rev1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   870
      Picture         =   "frmmain.frx":180C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   7050
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pTmp 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   330
      Picture         =   "frmmain.frx":1B4E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   7290
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pTmp2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   375
      Picture         =   "frmmain.frx":1E90
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   7035
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicBar2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00D1D8DB&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   2730
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   19
      Top             =   6300
      Width           =   4875
      Begin dmSpinFx.ButtonFx ButtonFx2 
         Height          =   330
         Left            =   855
         TabIndex        =   20
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Picture         =   "frmmain.frx":21D2
      End
      Begin dmSpinFx.ButtonFx ButtonFx1 
         Height          =   330
         Left            =   1275
         TabIndex        =   21
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
      End
      Begin dmSpinFx.ButtonFx ButtonFx3 
         Height          =   330
         Left            =   1620
         TabIndex        =   24
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
      End
      Begin dmSpinFx.ButtonFx ButtonFx4 
         Height          =   330
         Left            =   495
         TabIndex        =   40
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Picture         =   "frmmain.frx":2524
      End
      Begin dmSpinFx.ButtonFx ButtonFx5 
         Height          =   330
         Left            =   135
         TabIndex        =   41
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Picture         =   "frmmain.frx":2876
      End
      Begin dmSpinFx.ButtonFx ButtonFx6 
         Height          =   330
         Left            =   1965
         TabIndex        =   44
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
      End
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2460
         TabIndex        =   45
         Top             =   90
         Width           =   45
      End
      Begin VB.Image imgSpacer 
         Height          =   330
         Index           =   1
         Left            =   2325
         Top             =   30
         Width           =   15
      End
      Begin VB.Image imgSpacer 
         Height          =   330
         Index           =   0
         Left            =   1230
         Picture         =   "frmmain.frx":2BC8
         Top             =   30
         Width           =   15
      End
      Begin VB.Image Image1 
         Height          =   165
         Left            =   60
         Picture         =   "frmmain.frx":2C62
         Top             =   90
         Width           =   45
      End
   End
   Begin VB.PictureBox PicStatus 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   524
      TabIndex        =   17
      Top             =   6720
      Width           =   7860
      Begin dmSpinFx.dmHyperLink dmHyperLink1 
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   45
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   344
         HoverOut        =   -2147483630
         Caption         =   "#1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
      End
   End
   Begin VB.PictureBox PicA 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   45
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   15
      Top             =   1230
      Width           =   2235
      Begin VB.Label lblChrOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   30
         Width           =   1545
      End
   End
   Begin VB.PictureBox PicA 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   45
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   13
      Top             =   4200
      Width           =   2235
      Begin VB.Label lblChrOut 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character Output"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   30
         Width           =   1470
      End
   End
   Begin VB.ListBox lstPasswords 
      Height          =   4725
      IntegralHeight  =   0   'False
      ItemData        =   "frmmain.frx":2D28
      Left            =   2760
      List            =   "frmmain.frx":2D2A
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   1530
      Width           =   4920
   End
   Begin dmSpinFx.SpinFx SpinFx1 
      Height          =   315
      Left            =   135
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Max             =   2000
      Value           =   10
      Text            =   "10"
   End
   Begin VB.PictureBox PicDc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   7725
      Picture         =   "frmmain.frx":2D2C
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
      Top             =   3915
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picbar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   524
      TabIndex        =   1
      Top             =   615
      Width           =   7860
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "|"
         Height          =   195
         Left            =   1110
         TabIndex        =   47
         Top             =   90
         Width           =   30
      End
      Begin VB.Label lblButton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1230
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   90
         Width           =   255
      End
      Begin VB.Label lblSpacer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "|"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   90
         Width           =   30
      End
      Begin VB.Label lblButton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   615
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   90
         Width           =   420
      End
      Begin VB.Label lblButton 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   75
         MouseIcon       =   "frmmain.frx":2E1F
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   90
         Width           =   330
      End
   End
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7860
      TabIndex        =   0
      Top             =   0
      Width           =   7860
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This program is freeware"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5985
         TabIndex        =   4
         Top             =   60
         Width           =   1725
      End
      Begin VB.Label lblver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   3
         Top             =   240
         Width           =   1050
      End
      Begin VB.Image imgLogo 
         Height          =   510
         Left            =   60
         Picture         =   "frmmain.frx":2F71
         Top             =   60
         Width           =   4665
      End
   End
   Begin dmSpinFx.SpinFx SpinFx2 
      Height          =   315
      Left            =   135
      TabIndex        =   11
      Top             =   2595
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Max             =   500
      Value           =   10
      Text            =   "10"
   End
   Begin VB.Label lblSort 
      AutoSize        =   -1  'True
      Caption         =   "Password Sort Options"
      Height          =   195
      Left            =   135
      TabIndex        =   28
      Top             =   3330
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2595
      Index           =   2
      Left            =   60
      Top             =   1530
      Width           =   2580
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2205
      Index           =   1
      Left            =   60
      Top             =   4500
      Width           =   2580
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2205
      Index           =   0
      Left            =   45
      Top             =   4485
      Width           =   2580
   End
   Begin VB.Label lblPassLenLst 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Passwords to create:"
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   2355
      Width           =   2265
   End
   Begin VB.Label lblPassLen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Length:"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00C00000&
      X1              =   -30
      X2              =   735
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2595
      Index           =   3
      Left            =   45
      Top             =   1515
      Width           =   2580
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private cboTmp As String, cboIndex As Integer, TmpOld As String
Private Sort As Boolean, OptGenOption As Integer, bLoaded As Boolean


Sub GeneratePassword()
Dim X As Long, sLine As String, pLen As Integer

    lstPasswords.Clear
    pLen = Val(SpinFx1.Text)
    
    If pLen < 4 Then
        MsgBox "Your password must be at least 4 characters in length.", vbInformation, frmmain.Caption
        Exit Sub
    End If
    
    For X = 0 To Val(SpinFx2.Text) - 1
      
top:
        Select Case OptGenOption
            Case 0: sLine = GenPassword(26, 26, pLen)
            Case 1: sLine = GenPassword(0, 26, pLen)
            Case 2: sLine = GenPassword(52, 10, pLen)
            Case 3: sLine = GenPassword(0, 62, pLen)
            Case 4: sLine = GetExtPassword(pLen)
            Case 5: sLine = GetSpecialPassword(pLen)
            Case 6: sLine = GetPasswordAll(pLen)
            Case 7: sLine = GenPassword(52, 2, pLen)
        End Select
        
        DoEvents
        If SerachListBox(sLine, lstPasswords) <> -1 Then
        If ChkDup Then GoTo top:
    End If
    
        If cboIndex = 1 Then
            sLine = SortStr(sLine, True)
        End If
        
        If cboIndex = 2 Then
            sLine = SortStr(sLine, False)
        End If
        
        lstPasswords.AddItem sLine
    Next X
    
    
    lblGen.Caption = "Passwords Generated: " & X
    sLine = ""
End Sub

Sub LoadSettings()
Dim aTmp As String
    aTmp = GetSetting("DmPassGen", "config", "PassLen", 6)
    SpinFx1.Value = CInt(aTmp)
    aTmp = GetSetting("DmPassGen", "config", "NoOfPass", 6)
    SpinFx2.Value = CInt(aTmp)
    aTmp = GetSetting("DmPassGen", "config", "NoDups", 1)
    ChkDup.Value = CInt(aTmp)
    aTmp = GetSetting("DmPassGen", "config", "SortOrder", 0)
    cboSort.ListIndex = aTmp
    aTmp = GetSetting("DmPassGen", "config", "GenOption", 0)
    OptGen(CInt(aTmp)).Value = True
    aTmp = ""
    
End Sub

Sub InvertPassword(LstBox As ListBox)
Dim X As Long
    For X = 0 To LstBox.ListCount - 1
        LstBox.List(X) = Invert(LstBox.List(X))
    Next X
    X = 0
End Sub

Sub RevPassword(LstBox As ListBox)
Dim X As Long
    For X = 0 To LstBox.ListCount - 1
        LstBox.List(X) = Reserve(LstBox.List(X))
    Next X
    X = 0
End Sub

Function SortStr(lpStr As String, bSort As Boolean) As String
Dim X As Long, Y As Long, TmpStr() As String, Size As Long, sBuff As String

    Size = Len(lpStr) - 1
    ReDim Preserve TmpStr(Size)

    For X = 1 To Len(lpStr)
        TmpStr(X - 1) = Mid(lpStr, X, 1)
    Next X
    
    For X = 0 To Size
        For Y = X + 1 To Size
            If bSort Then
                If TmpStr(X) > TmpStr(Y) Then Swap TmpStr(X), TmpStr(Y)
            Else
                If TmpStr(X) < TmpStr(Y) Then Swap TmpStr(X), TmpStr(Y)
            End If
        Next Y
    Next X
    
    For X = 0 To Size
        sBuff = sBuff & TmpStr(X)
    Next
    
    Erase TmpStr
    SortStr = sBuff
    sBuff = ""
    X = 0: Y = 0: Size = 0
    
End Function

Sub FixSpin(SpinObj As SpinFx)
    If Not isalnum(SpinObj.Text) Then SpinObj.Text = ""
End Sub

Function SerachListBox(lpSerachFor As String, cboBox As ListBox) As Integer
Dim n As Integer

    SerachListBox = -1
    For n = 0 To cboBox.ListCount
        If cboBox.List(n) = lpSerachFor Then
            SerachListBox = n
            Exit For
        End If
    Next n
    
End Function

Sub SortLB(bSort As Boolean)
Dim X As Long, Y As Long, iSize As Long

    
    'Sort a Listbox
    iSize = (lstPasswords.ListCount) - 1
    If iSize = -1 Then Exit Sub
    
    ReDim Preserve TmpPasswords(iSize)

    For X = 0 To iSize
        TmpPasswords(X) = lstPasswords.List(X)
    Next X
    
    lstPasswords.Clear
    
    For X = 0 To iSize
        For Y = X + 1 To iSize
            If bSort Then
                If Left(TmpPasswords(X), 1) > Left(TmpPasswords(Y), 1) Then
                    Swap TmpPasswords(X), TmpPasswords(Y)
                End If
            Else
                If Left(TmpPasswords(X), 1) < Left(TmpPasswords(Y), 1) Then
                    Swap TmpPasswords(X), TmpPasswords(Y)
                End If
            End If
        Next Y
    Next X
    
    For X = 0 To iSize
        lstPasswords.AddItem TmpPasswords(X)
    Next X
    
    Erase TmpPasswords
    
    iSize = 0
    X = 0: Y = 0
End Sub

Private Sub UnLoadForm()
    SaveSetting "DmPassGen", "config", "PassLen", SpinFx1.Text
    SaveSetting "DmPassGen", "config", "NoOfPass", SpinFx2.Text
    SaveSetting "DmPassGen", "config", "NoDups", ChkDup.Value
    SaveSetting "DmPassGen", "config", "SortOrder", cboSort.ListIndex
    SaveSetting "DmPassGen", "config", "GenOption", OptGenOption
    cboTmp = ""
    TmpOld = ""
    Unload frmmain
End Sub

Sub DoHover(Index As Integer)
Dim X As Integer
    DoEvents
    For X = 0 To lblButton.Count - 1
        lblButton(X).FontUnderline = False
        lblButton(X).ForeColor = vbWhite
        Next X
    If Index = -1 Then Exit Sub
    
    lblButton(Index).ForeColor = &HE0E0E0
    lblButton(Index).FontUnderline = True
End Sub

Sub DrawBar()
Dim X As Long
Static Y As Integer

    For X = 0 To Picbar.ScaleWidth
        BitBlt Picbar.hdc, X, 0, PicDc.ScaleWidth, PicDc.ScaleHeight, PicDc.hdc, 0, 0, vbSrcCopy
    Next X
    
    For X = 0 To PicA(0).ScaleWidth
        Y = (Not Y)
        BitBlt PicA(Abs(Y)).hdc, X - 1, 0, PicDc.ScaleWidth, PicDc.ScaleHeight, PicDc.hdc, 0, 0, vbSrcCopy
    Next X
    
    For X = 0 To PicA(2).ScaleWidth
        BitBlt PicA(2).hdc, X - 1, 0, PicDc.ScaleWidth, PicDc.ScaleHeight, PicDc.hdc, 0, 0, vbSrcCopy
    Next
    
    PicStatus.Line (0, 0)-(PicStatus.ScaleWidth - 2, PicStatus.ScaleHeight - 1), &H808080, B
    PicBar2.Line (0, 0)-(PicBar2.ScaleWidth - 1, PicBar2.ScaleHeight - 1), &HA0A0A0, B

    Picbar.Refresh
    PicBar2.Refresh
    PicA(0).Refresh
    PicA(1).Refresh
    
    Set PicDc = Nothing

End Sub

Private Sub ButtonFx1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = "Sort passwords"
End Sub

Private Sub ButtonFx1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Sort = (Not Sort)
    
    If Not Sort Then
        Set ButtonFx1.Picture = pTmp.Picture
    Else
        Set ButtonFx1.Picture = pTmp2.Picture
    End If
    
    SortLB Sort
    
End Sub

Private Sub ButtonFx2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = "Copy passwords to Clipboad"
End Sub

Private Sub ButtonFx2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xCnt As Long, sLine As String

    If Button <> vbLeftButton Then Exit Sub
    
    'Copy selected list items to the clipboard
    If lstPasswords.ListCount = 0 Then Exit Sub
    For xCnt = 0 To lstPasswords.ListCount - 1
        If lstPasswords.Selected(xCnt) Then
            sLine = sLine & lstPasswords.List(xCnt) & vbCrLf
        End If
    Next xCnt
    MsgBox sLine
    Clipboard.Clear
    Clipboard.SetText sLine, vbCFText
    sLine = ""
    
End Sub

Private Sub ButtonFx3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = "Reserve passwords"
End Sub

Private Sub ButtonFx4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = "Export passwords List"
End Sub

Private Sub ButtonFx4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pSize As Long, Counter As Long

    If Button <> vbLeftButton Then Exit Sub
    If lstPasswords.ListCount = 0 Then
        MsgBox "There are no passwords to export yet" _
        & vbCrLf & "Please generate some passwords first.", vbInformation, "Export"
        Exit Sub
    Else
        pSize = (lstPasswords.ListCount) - 1
        
        Erase TmpPasswords
        ReDim Preserve TmpPasswords(pSize)
    
        For Counter = 0 To pSize
            TmpPasswords(Counter) = lstPasswords.List(Counter)
        Next Counter
    End If
    
    frmExport.Show vbModal, frmmain
End Sub

Private Sub ButtonFx5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = "Generate passwords"
End Sub

Private Sub ButtonFx5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Call GeneratePassword
End Sub

Private Sub ButtonFx6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = "Invert Passwords"
End Sub

Private Sub ButtonFx6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static iVert As Boolean

    If Button <> vbLeftButton Then Exit Sub
    iVert = Not iVert
    
    If iVert Then
        Set ButtonFx6.Picture = PicInv1.Picture
    Else
        Set ButtonFx6.Picture = PicInv2.Picture
    End If
    
    Call InvertPassword(lstPasswords)
End Sub

Private Sub cboSort_Change()
    cboSort.Text = cboTmp
End Sub

Private Sub cboSort_Click()
    cboTmp = cboSort.Text
    cboIndex = cboSort.ListIndex
End Sub

Private Sub ButtonFx3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static RevChar As Boolean

    If Button <> vbLeftButton Then Exit Sub
    RevChar = Not RevChar
    
    If RevChar Then
        Set ButtonFx3.Picture = Rev1.Picture
    Else
        Set ButtonFx3.Picture = rev2.Picture
    End If
    
    Call RevPassword(lstPasswords)
    
End Sub

Private Sub dmHyperLink1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    OpenSite frmmain.hwnd, "http://www.eraystudios.com"
    dmHyperLink1.ForeColor = dmHyperLink1.HoverOut
    dmHyperLink1.Font.Underline = False
End Sub

Private Sub Form_Load()
       
    DialogH.DlgHwnd = frmmain.hwnd
    DialogH.flags = 0
    DialogH.hInst = App.hInstance
    
    Sort = True
    
    For X = 1 To lblButton.Count - 1
        Set lblButton(X).MouseIcon = lblButton(0).MouseIcon
    Next
    Set dmHyperLink1.MouseIcon = lblButton(0).MouseIcon
    dmHyperLink1.Caption = "Copyright © 2004 - 2005 eRay Studios"
    imgSpacer(1).Picture = imgSpacer(0).Picture
    
    ButtonFx1_MouseUp vbLeftButton, 0, 0, 0
    ButtonFx3_MouseUp vbLeftButton, 0, 0, 0
    ButtonFx6_MouseUp vbLeftButton, 0, 0, 0
    
    cboSort.AddItem "No - Sorting"
    cboSort.AddItem "Sort - Ascending"
    cboSort.AddItem "Sort - Descending"
    cboSort.ListIndex = 0
    
    Call LoadSettings
    
    lblGen.Caption = "Passwords Generated:"
    
    bLoaded = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picbar_MouseMove Button, Shift, X, Y
    PicBar2_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UnLoadForm
End Sub

Private Sub Form_Resize()
    Picbar.Width = frmmain.ScaleWidth
    PicBar2.Width = lstPasswords.Width + 50
    
    ln1.X2 = frmmain.ScaleWidth
    Call DrawBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
    Set frmExport = Nothing
    Set DialogH = Nothing
    Set frmmain = Nothing
End Sub

Private Sub lblButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoHover Index
End Sub

Private Sub lblButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    Select Case Index
        Case 0: OpenSite frmmain.hwnd, FixPath(App.Path) & "help.chm"
        Case 1: frmabout.Show vbModal, frmmain
        Case 2: Call UnLoadForm
    End Select
    
End Sub

Private Sub OptGen_Click(Index As Integer)

    OptGenOption = Index
    If Val(SpinFx1.Text) < 4 Then Exit Sub
    
    If bLoaded Then
        Call GeneratePassword
        Exit Sub
    End If
    
End Sub

Private Sub Picbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoHover -1
End Sub

Private Sub PicBar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblTip.Caption = ""
End Sub

Private Sub SpinFx1_Change()
    FixSpin SpinFx1
End Sub

Private Sub SpinFx1_GotFocus()
    TmpOld = SpinFx1.Text
End Sub

Private Sub SpinFx1_LostFocus()
    If Len(SpinFx1.Text) = 0 Then SpinFx1.Text = TmpOld
End Sub

Private Sub SpinFx2_Change()
    FixSpin SpinFx2
End Sub

Private Sub SpinFx2_GotFocus()
    TmpOld = SpinFx2.Text
End Sub

Private Sub SpinFx2_LostFocus()
    If Len(SpinFx2.Text) = 0 Then SpinFx2.Text = TmpOld
End Sub
