VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   1485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   285
      Left            =   300
      TabIndex        =   6
      Top             =   1530
      Width           =   810
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fade Interval"
      Height          =   675
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   780
      Width           =   1305
      Begin VB.VScrollBar VScroll2 
         Height          =   330
         LargeChange     =   10
         Left            =   915
         Max             =   0
         Min             =   -1000
         SmallChange     =   10
         TabIndex        =   5
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         Height          =   330
         Left            =   240
         TabIndex        =   3
         Top             =   255
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Load Interval"
      Height          =   675
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   1305
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         LargeChange     =   100
         Left            =   930
         Max             =   0
         Min             =   -30000
         SmallChange     =   100
         TabIndex        =   4
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1000"
         Height          =   330
         Left            =   255
         TabIndex        =   2
         Top             =   225
         Width           =   660
      End
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Command1_Click()

    SaveSetting "AGR Slide Show", "Options", "Load Interval", VScroll1.value
    SaveSetting "AGR Slide Show", "Options", "Fade Interval", VScroll2.value
    Unload Me

End Sub

Private Sub Form_Load()

    VScroll1.value = -Form1.Timer1.Interval
    VScroll2.value = -Form1.Timer2.Interval

End Sub

Private Sub VScroll1_Change()

    Label1.Caption = Abs(VScroll1.value)
    Form1.Timer1.Interval = Abs(VScroll1.value)

End Sub

Private Sub VScroll2_Change()

    Label2.Caption = Abs(VScroll2.value)
    Form1.Timer2.Interval = Abs(VScroll2.value)

End Sub


