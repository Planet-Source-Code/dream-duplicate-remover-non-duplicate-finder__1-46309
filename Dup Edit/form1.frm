VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicate Editor"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7335
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5220
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   2461
            MinWidth        =   794
            Text            =   "Status:"
            TextSave        =   "Status:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:50 AM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edited List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   7095
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6855
      End
      Begin DuplicateEdit.lvButtons_H cmdClear3 
         Height          =   330
         Left            =   5640
         TabIndex        =   19
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Caption         =   "Clear"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin DuplicateEdit.lvButtons_H cmdSave3 
         Height          =   330
         Left            =   4440
         TabIndex        =   20
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label lblStatNum 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "       List1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2520
         Top             =   3240
      End
      Begin DuplicateEdit.lvButtons_H cmdSave1 
         Height          =   330
         Left            =   4440
         TabIndex        =   12
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         ItemData        =   "form1.frx":1272
         Left            =   120
         List            =   "form1.frx":1274
         TabIndex        =   7
         Top             =   2040
         Width           =   6855
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         ItemData        =   "form1.frx":1276
         Left            =   120
         List            =   "form1.frx":1278
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin DuplicateEdit.lvButtons_H cmdClear1 
         Height          =   330
         Left            =   5160
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Caption         =   "Clear"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin DuplicateEdit.lvButtons_H cmdAdd2 
         Height          =   330
         Left            =   5640
         TabIndex        =   14
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin DuplicateEdit.lvButtons_H cmdClear2 
         Height          =   330
         Left            =   4440
         TabIndex        =   15
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Caption         =   "Clear"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin DuplicateEdit.lvButtons_H cmdDupeRem1 
         Height          =   330
         Left            =   4440
         TabIndex        =   16
         ToolTipText     =   "This will remove all duplicates from List1"
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         Caption         =   "Remove Duplicates"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin DuplicateEdit.lvButtons_H cmdAdd1 
         Height          =   330
         Left            =   6000
         TabIndex        =   17
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         Caption         =   "Load"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin DuplicateEdit.lvButtons_H cmdEditBoth 
         Height          =   330
         Left            =   4440
         TabIndex        =   18
         ToolTipText     =   "This feature will find non duplicates comparing list2 against list1 and will list the non duplicates below!"
         Top             =   3360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         Caption         =   "Find Non Duplicates"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Total Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2640
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   7320
      X2              =   7320
      Y1              =   5520
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   0
      Y1              =   6120
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************* _
 Copyright © 2002-2003 Dream-Domain.net _
 ********************************* _
  NOTE: THIS HEADER MUST STAY INTACT.
' Terms of Agreement: _
  By using this code, you agree to the following terms... _
  1) You may use this code in your own programs (and may compile it into _
  a program and distribute it in compiled format for languages that allow _
  it) freely and with no charge. _
  2) You MAY NOT redistribute this code (for example to a web site) without _
  written permission from the original author. Failure to do so is a _
  violation of copyright laws. _
  3) You may link to this code from another website, but ONLY if it is not _
  wrapped in a frame. _
  4) You will abide by any additional copyright restrictions which the _
  author may have placed in the code or code's description. _
 **********************************
' Duplicate Remover _
  ----------------- _
  Example By Dream _
  Date: 6/20/03 10:15:37 AM _
  Email:  baddest_attitude@hotmail.com _
 ********************************** _
  Additional Terms of Agreement: _
  You MAY NOT Sell This Code _
  You MAY NOT Sell Any Program Containing This Code _
  You use this code knowing I hold no responsibilities for any results _
  occuring from the use and/or misuse of this code _
  If you make any improvements it would be nice if you would send me a copy. _
 ********************************** _
  Duplicate Remover _
 **********************************

Private nic(0 To 999999)       As String
Private nic2(0 To 999999)      As String
Private nic3(0 To 999999)      As String
Private a                      As Long
Private g                      As Long
Private z                      As Long
Private w                      As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

Private Sub cmdAdd1_Click()                     'list 1 load

  Dim sText As String
  Dim X     As Long

    On Error GoTo Err
    With CD1
        .DialogTitle = "Open List"
        .FileName = "List.txt"
        .Filter = "Text File |*.txt|"
        .ShowOpen
        Screen.MousePointer = 11
        X = FreeFile
        On Error Resume Next
        Open .FileName For Input As #X
        While Not EOF(X)
            Input #X, sText$
            If LenB(sText$) = 0 Then
                GoTo Err
             Else
                nic(z) = sText$
                z = z + 1
                DoEvents
            End If
        Wend
        Close #X
    End With
Err:
    Label5.Caption = z
    List1.AddItem CD1.FileName
    Screen.MousePointer = 1
    On Error GoTo 0
End Sub

Private Sub cmdAdd2_Click()                   '  list2  Load

  Dim sText As String
  Dim X     As Long

    On Error GoTo Err
    With CD1
        .DialogTitle = "Open List"
        .FileName = "List.txt"
        .Filter = "Text File |*.txt|"
        .ShowOpen
        Screen.MousePointer = 11
        X = FreeFile
        On Error Resume Next
        Open .FileName For Input As #X
        While Not EOF(X)
            Input #X, sText$
            If LenB(sText$) = 0 Then
                GoTo Err
             Else
                nic2(w) = sText$
                w = w + 1
                DoEvents
            End If
        Wend
        Close #X
    End With
Err:
    Label9.Caption = w
    List2.AddItem CD1.FileName
    Screen.MousePointer = 1
    On Error GoTo 0
End Sub

Private Sub cmdClear1_Click()

  Dim a As Long
    Screen.MousePointer = 11
    List1.Clear
    Label5.Caption = "0"
    For a = 0 To z
        nic(a) = vbNullString
    Next a
    z = 0
    Screen.MousePointer = 1
End Sub

Private Sub cmdClear2_Click()

  Dim a As Long
    Screen.MousePointer = 11
    List2.Clear
    Label9.Caption = "0"
    For a = 0 To w
        nic2(a) = vbNullString
    Next a
    w = 0
    Screen.MousePointer = 1

End Sub

Private Sub cmdClear3_Click()

  Dim a As Long
    Screen.MousePointer = 11
    Text1.Text = vbNullString
    lblStatNum.Caption = "0"
    For a = 0 To g
        nic3(a) = vbNullString
    Next a
    g = 0
    Screen.MousePointer = 1

End Sub

Private Sub cmdDupeRem1_Click()             'edit list 1

  Dim f As Long
  Dim b As Long

    If List1.ListCount < 1 Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    SB.Panels(2).Text = "Sorting..."
    Timer1 = True
    
   'Clear NonDuplicate array
    For f = 0 To g
        nic3(f) = vbNullString
    Next f
    g = 0
    
    'For every item in our array(our list of items)
    For a = 0 To z - 1
        
        'Check it against every item found after itself, no need to check it against
        'previous items
        For b = a + 1 To z - 1
            Select Case UCase(nic(b))
            '############################################################################
                'Replace any duplicate found with last item in array
                'Then removes one index from the array as we delete the duplicate item
                '
                'This loops until the new item from end of list is <> nic(a)
            'A Duplicate found, list the duplicate in seperate list
             Case Is = UCase(nic(a))
                nic3(g) = nic(b)
                
               'Replace with last item in array(our list of items)until
               'the item is a non duplicate
                Do Until UCase(nic(b)) <> UCase(nic(a))
                    nic(b) = nic(z - 1)
                   'Clear the last item in the array and reduce the array by 1
                    nic(z - 1) = vbNullString
                    z = z - 1
                   'Add item to duplicate list
                    g = g + 1
                    nic3(g) = nic(b)
                    DoEvents
                Loop
           '############################################################################
            End Select
            DoEvents
        Next b
        DoEvents
    Next a
    
    SB.Panels(2).Text = "Done"
    Label5.Caption = z
    lblStatNum.Caption = g
    Timer1 = False
    Screen.MousePointer = 1
    lblStat.Caption = "Duplicate List"

End Sub

Private Sub cmdEditBoth_Click()                 ' Edit both lists

  Dim k As Long
  Dim a As Long
  Dim b As Long
    If List1.ListCount < 1 Then
        Exit Sub
    End If
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    SB.Panels(2).Text = "Sorting.."
    For k = 0 To g
        nic3(g) = vbNullString
    Next k
    g = 0
    For a = 0 To w - 1
        For b = 0 To z - 1
            'If duplicate skip adding to new array
            If UCase(nic2(a)) = UCase(nic(b)) Then
                GoTo duplicate
            End If
            DoEvents
        Next b
            nic3(g) = nic2(a)
        g = g + 1
duplicate:
        DoEvents
    Next a
    SB.Panels(2).Text = "Done"
    Screen.MousePointer = 1
    lblStatNum.Caption = g
    lblStat.Caption = "Non Duplicate List"

End Sub

Private Sub cmdSave1_Click()                    'list1 save

  Dim Nbr As Long

    If Label5.Caption = "0" Then
        Exit Sub
    End If
    On Error GoTo Error_Killer
    With CD1
        .DialogTitle = "Save Edited List"
        .FileName = vbNullString
        .Filter = "Text File |*.txt|"
        .ShowSave
        Screen.MousePointer = 11
        On Error Resume Next
        Open .FileName For Output As #1
        For Nbr = 0 To z - 1
            Print #1, nic(Nbr)
        Next Nbr
        Close #1
    End With
    List1.AddItem CD1.FileName
    Screen.MousePointer = 1

Exit Sub

Error_Killer:
    Screen.MousePointer = 1
    MsgBox Error
    On Error GoTo 0
End Sub

Private Sub cmdSave3_Click()                    'list3    save

  Dim Nbr As Long

    If lblStatNum.Caption = "0" Then
        Exit Sub
    End If
    On Error GoTo Error_Killer
    With CD1
        .DialogTitle = "Save Edited List"
        .FileName = vbNullString
        .Filter = "Text File |*.txt|"
        .ShowSave
        Screen.MousePointer = 11
        On Error Resume Next
        Open .FileName For Output As #1
        For Nbr = 0 To g - 1
            Print #1, nic3(Nbr)
        Next Nbr
        Close #1
    End With
    Text1.Text = CD1.FileName
    Screen.MousePointer = 1

Exit Sub

Error_Killer:
    Screen.MousePointer = 1
    MsgBox Error
    On Error GoTo 0
End Sub

Private Sub Form_Keyup(KeyCode As Integer, _
                       Shift As Integer)

    Select Case KeyCode
     Case vbKeyF1   ' Form1.Command5_Click                'Show help form
     Case vbKeyEscape
        Unload Me                        'Unload program
    End Select

End Sub

Private Sub Form_QueryUnload(cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode = vbFormCode Then
        cancel = CBool(MsgBox("Do you wish to exit Duplicate Editor ?", vbYesNo Or vbInformation Or vbApplicationModal Or vbMsgBoxSetForeground Or vbSystemModal, Me.Caption) = vbNo)
    End If

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuReset_Click()

    cmdClear1_Click
    cmdClear2_Click
    cmdClear3_Click

End Sub

Private Sub Timer1_Timer()

    SB.Panels(2).Text = "Sorting.." & a
    Label5.Caption = z

End Sub

Public Function Win32Keyword(ByVal URL As String) As Long

    URL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)

End Function

''Private Sub label2002_Click()
'''<:-):WARNING: Unused Removed Control Code from a deleted control and has been commented out.
''Win32Keyword "www.dream-domain.net"
''End Sub
''Private Sub label2002_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''<:-):WARNING: Unused Removed Control Code from a deleted control and has been commented out.
''Label2002.ForeColor = &HFF0000
''End Sub
':) Roja's VB Code Fixer V1.0.97 (6/20/03 2:45:34 AM) 16 + 366 = 382 Lines Thanks Ulli for inspiration and lots of code.

