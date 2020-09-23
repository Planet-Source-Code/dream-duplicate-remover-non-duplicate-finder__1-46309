VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   15
   ClientTop       =   45
   ClientWidth     =   4395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox XpB 
      Height          =   135
      Left            =   3600
      MousePointer    =   99  'Custom
      ScaleHeight     =   75
      ScaleWidth      =   435
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      Height          =   2380
      Left            =   4440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmMain.frx":030A
      Top             =   90
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   200
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   615
      Left            =   90
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   3720
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   200
         Width           =   315
      End
      Begin VB.CommandButton Command3 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   3360
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   200
         Width           =   300
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Converter "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.Frame frmInfo 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   90
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
      Begin VB.Label lblinfo 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Click Here Anytime To Use The Calculator!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3975
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Converter.XpB time 
      Height          =   255
      Left            =   3260
      TabIndex        =   3
      Top             =   1240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Time"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   33023
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB temp 
      Height          =   255
      Left            =   1160
      TabIndex        =   0
      Top             =   1240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Temp"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   255
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB length 
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   1240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Length"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   8454016
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB weight 
      Height          =   255
      Left            =   2200
      TabIndex        =   2
      Top             =   1240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Weight"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   8454143
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB XpB1 
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   900
      Width           =   615
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "About"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   14737632
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB XpB2 
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   900
      Width           =   615
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "About"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   14737632
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB XpB3 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Pressure"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   12640511
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB XpB4 
      Height          =   255
      Left            =   1160
      TabIndex        =   18
      Top             =   900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Volume"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   12746315
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin Converter.XpB XpB5 
      Height          =   255
      Left            =   2200
      TabIndex        =   19
      Top             =   900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Area"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   2
      BackColor       =   32896
      SoundOver       =   ""
      SoundClick      =   ""
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFC0&
      X1              =   0
      X2              =   4320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   " "
      Height          =   2655
      Left            =   4320
      TabIndex        =   14
      Top             =   0
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   6
      X1              =   4320
      X2              =   4320
      Y1              =   2520
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   16
      X1              =   0
      X2              =   0
      Y1              =   2520
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   10
      X1              =   0
      X2              =   4320
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "By: Peter Hart  (c)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   1935
   End
   Begin VB.Menu cloesd 
      Caption         =   "Closed"
      Visible         =   0   'False
      Begin VB.Menu exit 
         Caption         =   "Exit?"
         Begin VB.Menu yes 
            Caption         =   "Yes"
         End
         Begin VB.Menu no 
            Caption         =   "No"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) 'everything above is for the Form Drag Code

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


 

Private Sub Command1_Click()
frmMain.PopupMenu Form2.closed
End Sub

Private Sub command2_click()
frmMain.PopupMenu Form2.FILE
End Sub

Private Sub command3_Click()
frmMain.WindowState = 1
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Lets you drag the form
   lblinfo.Visible = False
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
    Pause 0.09
 
  
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Lets you drag the form
  lblinfo.Visible = False
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
   Pause 0.09
 
 
End Sub


Private Sub lblinfo_Click()
Calculator.Show
End Sub
Private Sub label4_Click()
Calculator.Show
End Sub
Private Sub length_click()
 
frmlength.Show
End Sub


Private Sub length_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Between Metric And Imperial Units Of Length"
End Sub
Private Sub time_click()
frmtime.Show
End Sub


Private Sub time_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Seconds, Minutes, Hours, Days, Weeks, Months And years"
End Sub
Private Sub weight_click()
frmweight.Show
End Sub


Private Sub weight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Between Metric And Imperial Units Of Weight"
End Sub
Private Sub temp_Click()
frmtemp.Show
End Sub
Private Sub temp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Between Degrees Of Celcius And Farenheit"
End Sub

Private Sub Form_load()
XpB.Enabled = False
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me

End
End Sub

Private Sub XpB1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Click Here To Find Out More About Converter"
End Sub
Private Sub XpB2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Click Again To Close !!!!!"
End Sub




Private Sub XpB1_Click()
XpB2.Visible = True
   XpB1.Visible = False
     Width = 6420
End Sub

Private Sub XpB2_Click()

     XpB1.Visible = True
   XpB2.Visible = False
    Width = 4410
End Sub
Private Sub XpB3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Between Metric And Imperial Units Of Pressure"
End Sub
Private Sub XpB3_Click()
frmPres.Show
End Sub
Private Sub XpB4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Between Metric And Imperial Units Of Volume"
End Sub
Private Sub XpB4_Click()
frmVolu.Show
End Sub
Private Sub XpB5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Visible = True
lblinfo.Caption = "Converts Between Metric And Imperial Measurements Of Area"
End Sub
Private Sub XpB5_Click()
frmArea.Show
End Sub
