Attribute VB_Name = "Module1"
'ATTENTION ! This program uses SAVESETTING to Registry preferences.
'if you don't want to use your System Registry, please find and remove the lines.

'Credits to transitions procedures : Amiga Blitter
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=13409&lngWId=1

'To use this program as SCREEN SAVER Compile the Executable using the .SCR extension and
'save it into Windows/System directorie

Option Explicit
Public Screensaver As Boolean
Public Const APP_NAME = "Slide Show Screen_Saver"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Sub Main()
  
  Dim args As String
  
  Dim window_style As Long

    args = UCase$(Trim$(Command$))
      
    Select Case Mid$(args, 1, 2)
    
      Case "/C" ', ""   ' Display configuration dialog.
        Form1.Files_path_Click
        Form1.Show
      Case "/S" ', ""     ' Run as a screen saver.
        Screensaver = True
        Form1.Show
        Form1.Form_DblClick
      Case "/P"       ' Run in preview mode.
        Form1.Show
      Case Else
        If Not App.PrevInstance Then
            Form1.Show 0
            Exit Sub
        End If
    
        If FindWindow(vbNullString, APP_NAME) Then
            End
        End If
        
        Form2.Caption = APP_NAME
    
    End Select
  
End Sub


