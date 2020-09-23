VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export Password List"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4020
      TabIndex        =   6
      Top             =   1935
      Width           =   1155
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Export"
      Height          =   360
      Left            =   2760
      TabIndex        =   5
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export Options"
      Height          =   1680
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   5070
      Begin VB.OptionButton Op1 
         Caption         =   "Rich Text Format"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   1260
         Width           =   3585
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Internet HTML Document"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   960
         Width           =   3585
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Plain Text With Headings"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   660
         Width           =   3585
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Plain Text No Headings"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   345
         Width           =   3585
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dlFilter As String, ExportOp As Integer, FileExt As String

Private Sub cmdCancel_Click()
On Error Resume Next
    Erase TmpPasswords
    Unload frmExport
End Sub

Private Sub cmdOpen_Click()
Dim sFileExt As String, AbsFile As String, e_Pos As Integer

    With DialogH
        .DialogTitle = "Open"
        .Filter = dlFilter
        .ShowSave
        AbsFile = .FileName
        If Not .CancelError Then Exit Sub
    End With
  
    e_Pos = GetCharPos(AbsFile, ".", 1)
    
    If e_Pos = -1 Then
        AbsFile = AbsFile & "." & FileExt
    Else
        sFileExt = Trim(Mid(AbsFile, e_Pos, Len(AbsFile)))
        If sFileExt <> FileExt Then
            AbsFile = Mid(AbsFile, 1, e_Pos) & FileExt
        End If
    End If
    
    Select Case ExportOp
        Case 0
            ExportTxtPlain AbsFile, TmpPasswords
        Case 1
            ExportTxt AbsFile, TmpPasswords, frmmain.Caption
        Case 2
            ExportHtml AbsFile, TmpPasswords, frmmain.Caption
        Case 3
            ExportRtf AbsFile, TmpPasswords, frmmain.Caption
    End Select
    
    MsgBox "All passwords have been exported to:" & vbCrLf & AbsFile, vbInformation, "Export"
    
    AbsFile = ""
    cmdCancel_Click
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmExport = Nothing
End Sub

Private Sub Op1_Click(Index As Integer)
    ExportOp = Index
    
    Select Case ExportOp
        Case 0, 1
            dlFilter = "Text Files (*.txt)" & Chr(0) & "*.txt": FileExt = "txt"
        Case 2
            dlFilter = "HTML Documents (*.html)" & Chr(0) & "*.html": FileExt = "html"
        Case 3
            dlFilter = "Rich Text Format (*.rtf)" & Chr(0) & "*.rtf": FileExt = "rtf"
    End Select
    
End Sub

