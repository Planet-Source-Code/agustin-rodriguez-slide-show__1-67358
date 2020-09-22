VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3615
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Picture         =   "Form1.frx":164A
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   1215
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Timer Timer3 
      Left            =   285
      Top             =   5730
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   300
      Top             =   5100
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   2235
      TabIndex        =   1
      Top             =   4305
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   285
      Top             =   4500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   3225
      ScaleHeight     =   92
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   56
      TabIndex        =   0
      Top             =   4155
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Begin VB.Menu Files_path 
         Caption         =   "Set Directory"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Set_interval 
         Caption         =   "Set Interval"
      End
      Begin VB.Menu Aspect_rate 
         Caption         =   "Mantain aspect rate"
      End
      Begin VB.Menu Fullscreen 
         Caption         =   "Full Screen (Double Click)"
      End
   End
   Begin VB.Menu Transition_Type 
      Caption         =   "Transition Type"
      Begin VB.Menu opt 
         Caption         =   "Fade"
         Index           =   0
      End
      Begin VB.Menu opt 
         Caption         =   "Circle IN"
         Index           =   1
      End
      Begin VB.Menu opt 
         Caption         =   "Circle OUT"
         Index           =   2
      End
      Begin VB.Menu opt 
         Caption         =   "Implode"
         Index           =   3
      End
      Begin VB.Menu opt 
         Caption         =   "Hour Double"
         Index           =   4
      End
      Begin VB.Menu opt 
         Caption         =   "Close"
         Index           =   5
      End
      Begin VB.Menu opt 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu opt 
         Caption         =   "Random"
         Index           =   10
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ATTENTION ! This program uses SAVESETTING to Registry preferences.
'if you don't want to use your System Registry, please find and remove the lines.

'Credits to transitions procedures : Amiga Blitter
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=13409&lngWId=1

'To use this program as SCREEN SAVER Compile the Executable using the .SCR extension and
'save it into Windows/System directorie


Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_SHOWWINDOW As Long = &H40

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                          (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                          (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Private Const SWP_FRAMECHANGED As Integer = &H20
Private Const SWP_NOMOVE As Integer = &H2
Private Const SWP_NOZORDER As Integer = &H4
Private Const SWP_NOSIZE As Integer = &H1

Private Const WS_CAPTION = &HC00000
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hDC As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long

Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X _
                          As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As _
                          Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As _
                          Long, ByVal dwRop As Long) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                          (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                          (lpBrowseInfo As BROWSEINFO) As Long
       
Private Const BIF_RETURNONLYFSDIRS As Long = &H1

Private Trans_option As Integer
Private Random_Set As Integer
Private Path_target As String
Private Actual_picture As Integer
Private Pic_nr As Integer

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private BF As BLENDFUNCTION

Private Sub Aspect_rate_Click()

    Aspect_rate.Checked = Not Aspect_rate.Checked
    SaveSetting "AGR Slide Show", "Options", "Aspect", Aspect_rate.Checked
    
End Sub

Private Sub Exit_Click()

    Unload Me

End Sub

Public Sub Files_path_Click()

  Dim bi As BROWSEINFO
  Dim IDL As ITEMIDLIST
  Dim pidl As Long
  Dim r As Long
  Dim pos As Integer
  Dim spath As String
  
    bi.hOwner = Me.hwnd
       
    bi.pidlRoot = 0&
       
    bi.lpszTitle = "Select one Folder with Graphic files"
       
    bi.ulFlags = BIF_RETURNONLYFSDIRS
       
    pidl& = SHBrowseForFolder(bi)
       
    spath$ = Space$(512)
       
    r = SHGetPathFromIDList(ByVal pidl&, ByVal spath$)

    If r Then
        pos = InStr(spath$, Chr$(0))
        Path_target = Left$(spath$, pos - 1)
        If Right$(Path_target, 1) = "\" Then
            Path_target = Left$(Path_target, Len(Path_target) - 1)
        End If
      Else
        Exit Sub
    End If
    
    SaveSetting "AGR Slide Show", "Folder", "Path", Path_target
    
    Do_list
              
End Sub

Private Sub Form_Click()

    If Fullscreen.Checked Then
        Form_DblClick
    End If

End Sub

Public Sub Form_DblClick()

  Static X As Integer
  Dim rtn As Long
  Static w As Integer
  Static h As Integer
  Static l As Integer
  Static t As Integer

    Call fFlipBit(WS_CAPTION, X)
    X = X Xor 1

    If X Then
        Fullscreen.Checked = True
        WindowState = 0
        rtn = FindWindow("Shell_traywnd", "")
        'Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        w = Width
        h = Height
        l = Left
        t = Top
        Move 0, 0, Screen.Width, Screen.Height
    
      Else
        Fullscreen.Checked = False
        rtn = FindWindow("Shell_traywnd", "")
        'Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
        Move l, t, w, h
    End If

    Files.Visible = (X = 0)
    Options.Visible = (X = 0)
    Transition_Type.Visible = (X = 0)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Screensaver Then
        Unload Me
    End If
  
End Sub

Private Sub Form_Load()

    Randomize
    Path_target = GetSetting("AGR Slide Show", "Folder", "Path", App.Path)
    Trans_option = GetSetting("AGR Slide Show", "Transition", "Type", 0)
    Random_Set = (Trans_option = 10)
    Aspect_rate.Checked = GetSetting("AGR Slide Show", "Options", "Aspect", -1)

    Form2.VScroll1.value = GetSetting("AGR Slide Show", "Options", "Load Interval", -100)
    Form2.VScroll2.value = GetSetting("AGR Slide Show", "Options", "Fade Interval", -10)

    Do_list

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Static Xlast As Single, Ylast As Single
  Dim Xnow As Single
  Dim Ynow As Single

    If Not Screensaver Then
        Exit Sub
    End If

    Xnow = X
    Ynow = Y
    If Xlast = 0 And Ylast = 0 Then
        Xlast = Xnow
        Ylast = Ynow
    End If
    If (Xnow <> Xlast Or Ynow <> Ylast) Then
        Unload Me
    End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload Form2
    End

End Sub

Private Sub Fullscreen_Click()

    Form_DblClick

End Sub

Private Sub Opt_Click(Index As Integer)

    Trans_option = Index
    Random_Set = (Index = 10)
    SaveSetting "AGR Slide Show", "Transition", "Type", Trans_option
    
End Sub

Private Sub Set_interval_Click()

    Form2.Show 1

End Sub

Private Sub Timer1_Timer()

  Dim i As Double
  Dim w As Single
  Dim h As Single

    On Error GoTo erro
    
    Timer1.Enabled = False

    DoEvents

    If Pic_nr = List1.ListCount Then
        Pic_nr = 0
    End If
    
    'Caption = Str(Pic_nr + 1) & " /" & Str(List1.ListCount)
    
    Picture1.Picture = LoadPicture(Path_target & "\" & List1.List(Pic_nr))
    
    Picture2.BackColor = Picture1.Point(0, 0)
    Picture2.Width = ScaleWidth
    Picture2.Height = ScaleHeight
    
    Picture2.Cls
    
    If Aspect_rate.Checked Then
        w = Picture1.ScaleWidth
        h = Picture1.ScaleHeight
        Do While (h < ScaleHeight) Or (w < ScaleWidth)
            w = w + w / 100
            h = h + h / 100
            DoEvents
        Loop
        
        Do While (h > ScaleHeight) Or (w > ScaleWidth)
            w = w - w / 100
            h = h - h / 100
            DoEvents
        Loop
        
        Picture2.PaintPicture Picture1.Picture, (ScaleWidth - w) / 2, (ScaleHeight - h) / 2, w, h
        
      Else
        Picture2.PaintPicture Picture1.Picture, 0, 0, ScaleWidth, ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    End If
    
    Picture2.Refresh
    
    If List1.ListCount = 0 Then
        Exit Sub
    End If
    
    If Random_Set Then
        Trans_option = Int(Rnd * 6)
    End If
    
    Pic_nr = Pic_nr + 1
    Select Case Trans_option
      Case 0
        Timer2.Enabled = True
      Case 1, 2
        Do_trans_Circle
      Case 3
        Implode
      Case 4
        HourDblCB
      Case 5
        CloseOpt
      Case 10
    
    End Select
    
Exit_fade:

Exit Sub
    
erro:
    Resume Exit_fade

End Sub

Private Sub Timer2_Timer()

  Static i As Integer
    
    DoEvents
    AlphaBlend hDC, 0, 0, ScaleWidth, ScaleHeight, Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, i * 65536
    If i = 64 Then
        i = 0
        Timer2.Enabled = False
        Timer1.Enabled = True
    End If
    i = i + 2

End Sub

Private Sub Do_trans_Circle()

  Const PI = 3.1415
  Dim ray As Single
  Dim angle As Double
  Dim i As Integer
  Dim X As Single
  Dim Y As Single

    DoEvents
    ray = Sqr(Picture2.ScaleHeight ^ 2 + Picture2.ScaleWidth ^ 2) / 2

    If Trans_option = 1 Then
        For i = ray To 0 Step -2
            For angle = 0 To 2 * PI Step 0.01
                X = i * Cos(angle) + (ScaleWidth / 2)
                Y = i * Sin(angle) + (ScaleHeight / 2)
                BitBlt hDC, X, Y, 8, 8, Picture2.hDC, X, Y, vbSrcCopy
            Next angle
        Next i

      Else
    
        For i = 0 To ray Step 2
            For angle = 0 To 2 * PI Step 0.01
                X = i * Cos(angle) + (ScaleWidth / 2)
                Y = i * Sin(angle) + (ScaleHeight / 2)
                BitBlt hDC, X, Y, 8, 8, Picture2.hDC, X, Y, vbSrcCopy
            Next angle
        Next i
    End If

    Timer1.Enabled = True

End Sub

Private Sub Implode()

  Const PI = 3.1415
  Dim i As Integer
  Dim ray As Single
  Dim angle As Double
  Dim X As Single
  Dim Y As Single
    
    On Error GoTo erro
    
    DoEvents
    ray = Sqr(Picture2.ScaleHeight ^ 2 + Picture2.ScaleWidth ^ 2) / 2
    For i = ray To 0 Step -5
        For angle = 0 To 5 * PI Step 0.01
            X = i * Tan(angle) + (ScaleWidth / 2)
            Y = i * Cos(angle) + (ScaleHeight / 2)
            BitBlt hDC, X, Y, 16, 16, Picture2.hDC, X, Y, vbSrcCopy
            DoEvents
        Next angle
    Next i
sair:
    Timer1.Enabled = True

Exit Sub

erro:
    Resume sair

End Sub

Private Sub HourDblCB()

  Const PI = 3.1415
  Dim ray As Single
  Dim angle As Double
  Dim a As Double
  Dim b As Double
  Dim c As Double
  Dim X As Double
  Dim Y As Double

    On Error GoTo erro

    DoEvents

    For angle = 0 To 2 * PI Step 0.01
        a = Tan(angle)
        b = Cos(angle)
        c = Sin(angle)
        If Abs(a * (ScaleWidth / 2)) < (ScaleHeight / 2) Then
            For X = -0.5 * (1 + Sgn(b)) * (ScaleWidth / 2) To 0.5 * (1 + Sgn(b)) * (ScaleWidth / 2) Step Sgn(b)
                BitBlt hDC, (ScaleWidth / 2) + X, (ScaleHeight / 2) + a * X, 8, 8, Picture2.hDC, _
                       (ScaleWidth / 2) + X, (ScaleHeight / 2) + a * X, vbSrcCopy
                DoEvents
            Next X
          Else
            For Y = -1 * (1 + Sgn(c)) * (ScaleWidth / 2) To 1 * (1 + Sgn(c)) * (ScaleWidth / 2) Step Sgn(c)
                BitBlt hDC, (ScaleWidth / 2) + Y / a, (ScaleHeight / 2) + Y, 8, 8, Picture2.hDC, _
                       (ScaleWidth / 2) + Y / a, (ScaleHeight / 2) + Y, vbSrcCopy
                DoEvents
            Next Y
        End If
    Next angle
sair:
    Timer1.Enabled = True
    
Exit Sub

erro:
    Resume sair

End Sub

Private Sub CloseOpt()

  Dim ImgX As Integer
  Dim ImgY As Integer
  Dim NumLoop As Integer
  Dim HalfHeight As Integer
  Dim i As Integer
  Dim X As Double
  Dim Y As Double

    On Error GoTo erro

    ImgX = ScaleWidth
    ImgY = ScaleHeight
    HalfHeight = ImgY / 2

    For i = 0 To HalfHeight + 5 Step 5
        Y = i
        For X = i To ImgX - i
            BitBlt hDC, X, Y, 5, 5, Picture2.hDC, X, Y, vbSrcCopy
            DoEvents
        Next X
    
        Wait (5)
    
        X = ImgX - i
        For Y = i To ImgY - i
            BitBlt hDC, X, Y, 5, 5, Picture2.hDC, X, Y, vbSrcCopy
            DoEvents
        Next Y
        Wait (5)
        Y = ImgY - i
        For X = ImgX - i To i Step -5
            BitBlt hDC, X, Y, 5, 5, Picture2.hDC, X, Y, vbSrcCopy
            DoEvents
        Next X
        Wait (5)
        X = i
        For Y = ImgY - i To i Step -5
            BitBlt hDC&, X, Y, 5, 5, Picture2.hDC, X, Y, vbSrcCopy
            DoEvents
        Next Y
        Wait (5)
        DoEvents
    Next i

sair:
    Timer1.Enabled = True

Exit Sub

erro:
    Resume sair

End Sub

Private Function Wait(ByVal TimeToWait As Long)

  Dim EndTime As Long

    EndTime = GetTickCount + TimeToWait

    Do Until GetTickCount > EndTime
        DoEvents
    Loop

End Function

Private Function fFlipBit(ByVal Bit As Long, ByVal value As Boolean) As Boolean

  Dim lStyle As Long
   
    lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
   
    If value Then
        lStyle = lStyle Or Bit
      Else
        lStyle = lStyle And Not Bit
    End If
    Call SetWindowLong(Me.hwnd, GWL_STYLE, lStyle)
    Call pRedraw
   
    fFlipBit = (lStyle = GetWindowLong(Me.hwnd, GWL_STYLE))

End Function

Private Sub pRedraw()

  
  
  Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

    Call SetWindowPos(Me.hwnd, 0, 0, 0, 0, 0, swpFlags)

End Sub

Private Sub Do_list()

  Dim X As String

    List1.Clear
    Timer1.Enabled = False
    Pic_nr = 0
    X = Dir$(Path_target & "\*.jpg")
    Do While X <> ""
        List1.AddItem X
        DoEvents
        X = Dir
    Loop
           
    X = Dir$(Path_target & "\*.bmp")
    Do While X <> ""
        List1.AddItem X
        DoEvents
        X = Dir
    Loop
    
    X = Dir$(Path_target & "\*.gif")
    Do While X <> ""
        List1.AddItem X
        DoEvents
        X = Dir
    Loop
    
    Actual_picture = 0
    If List1.ListCount Then
        Timer1.Enabled = True
        Caption = ""
      Else
        Timer1.Enabled = False
        Caption = "Folder without picture"
    End If

End Sub

