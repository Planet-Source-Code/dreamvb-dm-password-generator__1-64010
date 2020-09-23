Attribute VB_Name = "modTools"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public DialogH As New CDialog
Public TmpPasswords() As String

Public Function GetCharPos(lpStr As String, lFindChar As String, sPosition As Integer) As Integer
Dim e_Pos As Integer, X

    e_Pos = -1
    For X = 1 To Len(lpStr)
        If LCase(Mid(lpStr, X, 1)) = LCase(lFindChar) Then
            e_Pos = X
            If sPosition = 0 Then ' find first
                Exit For
            End If
        End If
    Next
    
    X = 0
    GetCharPos = e_Pos
    e_Pos = 0
    
End Function

Function GetFilename(lpFile As String) As String
Dim e_Pos As Integer
    e_Pos = GetCharPos(lpFile, "\", 1)
    If e_Pos = -1 Then
        GetFilename = lpFile
    Else
        GetFilename = Trim(Mid(lpFile, e_Pos + 1, Len(lpFile)))
    End If
    e_Pos = 0
End Function

Function GetFileExt(lpFile As String) As String
Dim e_Pos As Integer
    e_Pos = GetCharPos(lpFile, ".", 1)
    If e_Pos = -1 Then
        GetFileExt = lpFile
        Exit Function
    Else
        GetFileExt = Trim(Mid(lpFile, e_Pos, Len(lpFile)))
        e_Pos = 0
    End If
End Function

Function isalnum(lpStr As String) As Boolean
Dim c As String, e_Pos As Integer

    For X = 1 To Len(lpStr)
        c = Mid(lpStr, X, 1)
        Select Case c
            Case "0" To "9": e_Pos = 1
            Case Else
                e_Pos = 0
                Exit For
        End Select
    Next X
    
    isalnum = e_Pos
    X = 0
    
End Function
Public Sub Swap(A, B)
Dim t
    'Swaps two values
    t = B
    B = A
    A = t
End Sub

Public Sub OpenSite(mHwnd As Long, lpUrl As String)
    ShellExecute mHwnd, "open", lpUrl, vbNullString, vbNullString, 1
End Sub

Function FixPath(lzPath As String) As String
   If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function Reserve(lpText As String)
Dim X As Integer, ch As String
    For X = Len(lpText) To 1 Step -1
        ch = ch & Mid$(lpText, X, 1)
    Next X
    
    Reserve = ch
    ch = ""
    X = 0
    
End Function

Function DeleteFileFx(lpFileName As String)
On Error Resume Next
    SetAttr lpFileName, vbNormal
    SaveFile lpFileName, ""
    Kill lpFileName
End Function

Function Invert(lpText As String)
Dim X As Integer, sBuff As String, ch As String
    For X = 1 To Len(lpText)
        ch = Mid$(lpText, X, 1)
        
        If ch = UCase(ch) Then
            ch = LCase(ch)
        Else
            ch = UCase(ch)
        End If
        
        sBuff = sBuff & ch
    Next X
    
    X = 0
    ch = ""
    Invert = sBuff
    sBuff = ""
End Function

Function GenPassword(Hi As Integer, lo As Integer, Length As Integer) As String
Dim X As Integer, s As String
    'Password generator 1
    Const Chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    For X = 1 To Length
        Randomize
        s = s & Mid(Chars, (Hi + Int(Rnd * lo) + 1), 1)
    Next X
    
    X = 0
    GenPassword = s
    
End Function

Function GetExtPassword(Length As Integer) As String
Dim X As Integer, s As String
    'Password generator 2
    For X = 1 To Length
        Randomize
        s = s & Chr((127 + Int(Rnd * 127) + 1))
    Next X
    
    GetExtPassword = s
    s = ""
End Function

Function GetSpecialPassword(Length As Integer) As String
Dim X As Integer, s As String
    'Password generator 3
    Const sp_chars = "!""#$%&'()*+,-./:;<=>?@[\]^_`{}~|"
    
    For X = 1 To Length
        Randomize
        s = s & Mid(sp_chars, (Int(Rnd * 32) + 1), 1)
    Next X
    
    X = 0
    GetSpecialPassword = s
    s = ""
    
End Function

Function GetPasswordAll(Length As Integer) As String
Dim pTmp As String
    pTmp = GetExtPassword(127) & "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!""#$%&'()*+,-./:;<=>?@[\]^_`{}~|"
    
    For X = 1 To Length
        Randomize
        s = s & Mid(pTmp, (Int(Rnd * Len(pTmp)) + 1), 1)
    Next X
    
    X = 0
    GetPasswordAll = s
    s = ""
    
End Function

Sub SaveFile(lpFile As String, lpData As String)
Dim fp As Long
    fp = FreeFile
    Open lpFile For Output As #fp
        Print #fp, lpData;
    Close #fp
End Sub

Sub ExportRtf(lExportFile As String, mList() As String, lTemp As String)
Dim iSize As Long, X As Long, nEnd As String
Dim RtfHead As String, sList As String, sMsg As String

    iSize = UBound(mList)
    
    If iSize = 0 Then
        sMsg = "password generated."
    Else
        sMsg = "passwords generated."
    End If
    
    RtfHead = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033\deflangfe1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}}" & vbCrLf
    RtfHead = RtfHead & "{\colortbl ;\red0\green0\blue255;}" & vbCrLf
    RtfHead = RtfHead & "{\*\generator " & lTemp & ";}\viewkind4\uc1\pard\nowidctlpar\cf1\b\f0\fs32 Generated Password List\cf0\par" & vbCrLf
    RtfHead = RtfHead & "\fs16 " & iSize + 1 & " " & sMsg & "\par" & vbCrLf
    RtfHead = RtfHead & "\par\fs20" & vbCrLf
    
    For X = 0 To iSize
        If X = iSize Then nEnd = "\fs32"
        sList = sList & "\fs20 " & (X + 1) & ".\tab " & mList(X) & nEnd & "\par" & vbCrLf
    Next X
    
    sList = sList & "\b0\par" & vbCrLf & "}" & vbCrLf
    
    SaveFile lExportFile, RtfHead & sList
    sMsg = ""
    RtfHead = ""
    sList = ""
    nEnd = ""
    iSize = 0
    Erase mList
    
End Sub

Sub ExportHtml(lExportFile As String, mList() As String, lTemp As String)
Dim iSize As Long, X As Long
Dim htmHead As String, sList As String, sMsg As String

    iSize = UBound(mList)
    
    If iSize = 0 Then
        sMsg = "password generated."
    Else
        sMsg = "passwords generated."
    End If
    
    htmHead = "<html>" & vbCrLf
    htmHead = htmHead & "<head>" & vbCrLf
    htmHead = htmHead & "<title>Generated Password List</title>" & vbCrLf
    htmHead = htmHead & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf
    htmHead = htmHead & "</head>" & vbCrLf
    htmHead = htmHead & vbCrLf
    htmHead = htmHead & "<body bgcolor=""#FFFFFF"" text=""#000000"">" & vbCrLf
    htmHead = htmHead & "<!--" & lTemp & " -->" & vbCrLf
    htmHead = htmHead & "<p><b><font face=""Arial, Helvetica, sans-serif"" size=""4"" color=""#0000FF"">Generated" & vbCrLf
    htmHead = htmHead & "Password List</font></b><br>" & vbCrLf
    
    htmHead = htmHead & "  <font size=""1"" face=""Arial, Helvetica, sans-serif""><b>" & (iSize + 1) & " " & sMsg & "</b></font>" & vbCrLf
    htmHead = htmHead & "<ol>" & vbCrLf
    
    For X = 0 To iSize
       sList = sList & "<li><font face=""Arial, Helvetica, sans-serif"" size=""2""><b>" & mList(X) & "</b></font></li>" & vbCrLf
    Next X
    
    sList = sList & "</ol>" & vbCrLf
    sList = sList & "</body>" & vbCrLf
    sList = sList & "</html>" & vbCrLf
    
    SaveFile lExportFile, htmHead & sList
    sMsg = ""
    htmHead = ""
    sList = ""
    iSize = 0
    Erase mList
    
End Sub

Sub ExportTxt(lExportFile As String, mList() As String, lTemp As String)
Dim iSize As Long, X As Long
Dim TxtHead As String, sList As String, sMsg As String

    iSize = UBound(mList)
    
    If iSize = 0 Then
        sMsg = "password generated."
    Else
        sMsg = "passwords generated."
    End If
    
    TxtHead = "Generated Password List" & vbCrLf
    TxtHead = TxtHead & "- - - - - - - - - - - - - - - - - - -" & vbCrLf
    TxtHead = TxtHead & iSize + 1 & " " & sMsg & vbCrLf & vbCrLf
    
    For X = 0 To iSize
        sList = sList & (X + 1) & "." & vbTab & mList(X) & vbCrLf
    Next X
    
    sList = sList & vbCrLf & vbCrLf & lTemp & vbCrLf

    SaveFile lExportFile, TxtHead & sList
    sMsg = ""
    TxtHead = ""
    sList = ""
    iSize = 0
    Erase mList
    
End Sub

Sub ExportTxtPlain(lExportFile As String, mList() As String)
Dim iSize As Long, X As Long
Dim sList As String

    iSize = UBound(mList)
    
    For X = 0 To iSize
        sList = sList & mList(X) & vbCrLf
    Next X

    SaveFile lExportFile, TxtHead & sList
    sList = ""
    iSize = 0
    Erase mList
    
End Sub
