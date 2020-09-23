VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computerized Examination version 1.0"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   900
      Top             =   660
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   6705
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11324
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDesign 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6735
      Left            =   0
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   632
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   9540
      Begin VB.CommandButton cmdDirection 
         BackColor       =   &H00BFF0EF&
         Caption         =   "&Direction"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton cmdNextCategory 
         BackColor       =   &H00BFF0EF&
         Caption         =   "Next &Category"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "&Next"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1335
      End
      Begin VB.PictureBox picTray 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00BFF0EF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   2160
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   229
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2640
         Width           =   3435
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ok"
            Default         =   -1  'True
            Height          =   375
            Left            =   420
            TabIndex        =   0
            Top             =   1080
            WhatsThisHelpID =   9000
            Width           =   975
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00C0C0C0&
            Cancel          =   -1  'True
            Caption         =   "Exit"
            Height          =   375
            Left            =   2220
            TabIndex        =   1
            Top             =   1080
            WhatsThisHelpID =   9000
            Width           =   975
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3120
            TabIndex        =   31
            Top             =   60
            Width           =   105
         End
         Begin VB.Label lblMess 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Would you like to start now?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   540
            TabIndex        =   30
            Top             =   480
            Width           =   2460
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   29
            Top             =   60
            Width           =   105
         End
      End
      Begin VB.PictureBox picDummy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000CCFF&
         Height          =   315
         Left            =   1860
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   660
         Width           =   255
      End
      Begin VB.Timer tmrTimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1320
         Top             =   660
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "N&ext Flag"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2700
         Width           =   1335
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "&Last Flag"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2100
         Width           =   1335
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "&Flag"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "Next Unanswered"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   6
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "Last Unaswered"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdPos 
         BackColor       =   &H00BFF0EF&
         Caption         =   "&Previous"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox picTime 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   6120
         Width           =   2415
      End
      Begin VB.PictureBox picCompleted 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   161
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6120
         Width           =   2415
      End
      Begin MSComctlLib.ImageCombo imcListAns 
         Height          =   420
         Left            =   240
         TabIndex        =   22
         Tag             =   "MarkButton"
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   15006974
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         ImageList       =   "imlImageList"
      End
      Begin MSComctlLib.ImageList imlImageList 
         Left            =   240
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   15006974
         ImageWidth      =   22
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CE314
               Key             =   "imgLightOff"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CE768
               Key             =   "imgLightOn"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pgbMinute 
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   6300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pgbSecond 
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   6060
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   59
         Scrolling       =   1
      End
      Begin prjExamination.ChoiceButton cbOption 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   2280
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
      End
      Begin prjExamination.ChoiceButton cbOption 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   3180
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
         Selection       =   1
      End
      Begin prjExamination.ChoiceButton cbOption 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   4020
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
         Selection       =   2
      End
      Begin prjExamination.ChoiceButton cbOption 
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   4860
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   661
         Selection       =   3
      End
      Begin VB.Label lblQuestion 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1395
         Left            =   1380
         TabIndex        =   20
         Top             =   60
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label lblOption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   1080
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.Label lblOption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   1080
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.Label lblOption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   3960
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.Label lblOption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   1080
         TabIndex        =   15
         Top             =   4800
         Visible         =   0   'False
         Width           =   5835
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' i want to add some more but the file is to large. It's
' hard for me to upload.

Dim CurPos     As Integer
Dim FlagWidth  As Long
Dim FlagHeight As Long
Dim TotalAns   As Long

Dim hh        As Byte
Dim mm        As Byte
Dim ss        As Byte
Dim TotalMins As Integer

Dim FlagPict As New StdPicture

Private Sub cbOption_Click(Index As Integer)
    Dim i As Integer
    
    For i = 0 To cbOption.Count - 1
        If Index <> i Then
            cbOption(i).Press = False
            lblOption(i).BackStyle = 0
        Else
            lblOption(Index).BackStyle = IIf(cbOption(Index).Press, 1, 0)
            lblOption(Index).BackColor = &H80FFFF
            imcListAns.ComboItems(CurPos).Image = _
                IIf(cbOption(Index).Press, "imgLightOn", "imgLightOff")
                
            If cbOption(Index).Press Then
                MyAnswer(CurPos).Answer = DataInfo(DataInfo(CurPos).ItemID).OptionID(Index + 1)
                MyAnswer(CurPos).Selected = Index
            Else
                MyAnswer(CurPos).Answer = 0
            End If
        End If
    Next
    
    TotalAns = 0
    For i = 1 To MyAnswer.Count
        If MyAnswer(i).Answer <> 0 Then
            TotalAns = TotalAns + 1
        End If
    Next i
    AnswerCompleted "Completed at - " & CStr(TotalAns) & "/" & MyExam.TotalItems, vbBlack, -1
    Debug.Print DataInfo(DataInfo(CurPos).ItemID).OptionID(Index + 1)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    
    StartCategory
    picTray.Visible = False
    picDummy.Visible = False
    lblQuestion.Visible = True
    imcListAns.Visible = True
    For i = 0 To 3
        cbOption(i).Visible = True
        lblOption(i).Visible = True
    Next
    For i = cmdPos.LBound To cmdPos.UBound
        cmdPos(i).Enabled = True
    Next i
    cmdNextCategory.Enabled = True
    cmdDirection.Enabled = True
    cbOption(0).SetFocus
End Sub

Private Sub cmdPos_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
    Case Is = 0 ' Next
        CurPos = CurPos + 1
        If CurPos > MyExam.TotalItems Then
            CurPos = MyExam.TotalItems
        End If
    Case Is = 1 ' Previous
        CurPos = CurPos - 1
        If CurPos < 1 Then CurPos = 1
    Case Is = 2 ' Flag
        MyAnswer(CurPos).Flag = Not MyAnswer(CurPos).Flag
        If MyAnswer(CurPos).Flag Then
            DisplayFlaggedPicture
        Else
            picDesign.Cls
        End If
        Exit Sub
    Case Is = 3 ' Last Flag
        For i = CurPos - 1 To 1 Step -1
            If MyAnswer(i).Flag Then
                CurPos = i
                Exit For
            End If
        Next i
    Case Is = 4 ' Next flag
        For i = CurPos + 1 To MyAnswer.Count
            If MyAnswer(i).Flag Then
                CurPos = i
                Exit For
            End If
        Next i
    Case Is = 5 ' Last Unanswered
        For i = CurPos - 1 To 1 Step -1
            If MyAnswer(i).Answer = 0 Then
                CurPos = i
                Exit For
            End If
        Next i
    Case Is = 6 ' Next Unanswered
        For i = CurPos + 1 To MyAnswer.Count
            If MyAnswer(i).Answer = 0 Then
                CurPos = i
                Exit For
            End If
        Next i
    End Select
        
    UpdateQuestion
End Sub

Private Sub cmdDirection_Click()
    frmDirection.Show vbModal, Me
End Sub

Private Sub cmdNextCategory_Click()
    Dim Mess As String, ret As Boolean
     
     ret = False
    If TotalAns <> MyExam.TotalItems Then
        Mess = MyExam.TotalItems - TotalAns & " questions unanswered!" _
            & Chr(13) & "Next Category?"
        If MsgBox(Mess, vbYesNo Or vbInformation, "Examination") = vbYes Then
            ret = True
        End If
    Else
        Mess = "Next Category?"
        If MsgBox(Mess, vbOKCancel Or vbInformation, "Examination") = vbOK Then
            ret = True
        End If
    End If
    
    If ret Then
        MyExam.ComputeScore
        StartCategory
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = AscW(UCase$(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyA Then
        cbOption(0).Press = Not cbOption(0).Press
        cbOption_Click 0
        cbOption(0).SetFocus
    ElseIf KeyAscii = vbKeyB Then
        cbOption(1).Press = Not cbOption(1).Press
        cbOption_Click 1
        cbOption(1).SetFocus
    ElseIf KeyAscii = vbKeyC Then
        cbOption(2).Press = Not cbOption(2).Press
        cbOption_Click 2
        cbOption(2).SetFocus
    ElseIf KeyAscii = vbKeyD Then
        cbOption(3).Press = Not cbOption(3).Press
        cbOption_Click 3
        cbOption(3).SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim cx As Integer, cy As Integer, tRECT As RECT, i As Integer
    Dim lhDC As Long, PicWidth As Long, PicHeight As Long
    On Error Resume Next
    
    lblDate.Caption = Date
    lblTime.Caption = Time
        
    For i = cmdPos.LBound To cmdPos.UBound
        Set cmdPos(i).Picture = LoadResPicture(998, vbResBitmap)
        Set cmdPos(i).DisabledPicture = LoadResPicture(998, vbResBitmap)
    Next i
    Set cmdNextCategory.Picture = LoadResPicture(999, vbResBitmap)
    Set cmdNextCategory.DisabledPicture = LoadResPicture(999, vbResBitmap)
    Set cmdDirection.Picture = LoadResPicture(999, vbResBitmap)
    Set cmdDirection.DisabledPicture = LoadResPicture(999, vbResBitmap)
    
    GetClientRect picTray.hwnd, tRECT
    DrawEdge picTray.hdc, tRECT, &H5, BF_RECT
    DrawLine picTray.hdc, 15, 60, tRECT.Right - 15, 60, 1, &H8000000F
    DrawLine picTray.hdc, 16, 61, tRECT.Right - 14, 61, 1, vbWhite
    
    Dim tBM As BITMAP, Temp As IPicture
    Set Temp = LoadPicture(ImagePath & "picnbtop.bmp")
    GetObjectAPI Temp.Handle, Len(tBM), tBM
    picDesign.PaintPicture Temp, 0, 0, tBM.bmWidth, tBM.bmHeight
    Set Temp = LoadPicture(ImagePath & "picnbbody.bmp")
    GetObjectAPI Temp.Handle, Len(tBM), tBM
    picDummy.Move 11, 129, 469, 252
    GetClientRect picDummy.hwnd, tRECT
    For cy = 0 To tRECT.Bottom / tBM.bmHeight
        For cx = 0 To tRECT.Right / tBM.bmWidth
            picDummy.PaintPicture Temp, cx * tBM.bmWidth, cy * tBM.bmHeight, tBM.bmWidth, tBM.bmHeight
        Next cx
    Next cy
    If MyExam.Title <> "" Then
        DrawText picDummy.hdc, MyExam.Title, Len(MyExam.Title), tRECT, DT_CENTER Or DT_WORDBREAK
    End If
    
    Set FlagPict = LoadPicture(ImagePath & "flag.bmp ")
    
    AnswerCompleted "Completed at - 0/0", vbWhite, &H8000000C
    DisplayTimer "Time : 00:00:00", vbWhite, &H8000000C
    
    sbStatusBar.Panels("Date").text = Date
    sbStatusBar.Panels("Time").text = Time
    tmrTime.Enabled = True
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    Dim resp As Integer
    
    resp = MsgBox("Are you sure you want to exit?", vbYesNo Or vbQuestion, "Exit")
    
    If resp = vbYes Then
        MyExam.CleanUp
    Else
        Cancel = True
    End If
End Sub

Private Sub imcListAns_Click()
    CurPos = Val(imcListAns.SelectedItem.text)
    UpdateQuestion
End Sub

Private Sub UpdateQuestion()
    MyExam.Question DataInfo(CurPos).ItemID
    imcListAns.ComboItems(CurPos).Selected = True
    
    Dim i As Integer
    
    For i = 0 To cbOption.Count - 1
        picDesign.Cls   ' clear flagged question
        cbOption(i).Press = False
        lblOption(i).BackStyle = 0
    Next i
    
    If MyAnswer(CurPos).Answer <> 0 Then
        For i = 0 To cbOption.Count - 1
            If i <> MyAnswer(CurPos).Answer - 1 Then
                Dim n As Integer
                
                n = MyAnswer(CurPos).Selected
                cbOption(n).Press = True
                lblOption(n).BackStyle = 1
                lblOption(n).BackColor = &H80FFFF
                Exit For
            End If
        Next i
    End If
    
    If MyAnswer(CurPos).Flag Then DisplayFlaggedPicture
End Sub

Private Sub mnuAbout_Click()
    MsgBox "email:humsafar_ak@yahoo.com", vbInformation, "About Me"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tmrTime_Timer()
    If lblDate.Visible Then
        lblDate.Caption = Date
        lblTime.Caption = Time
    End If
    
    sbStatusBar.Panels("Date").text = Date
    sbStatusBar.Panels("Time").text = Time
End Sub

Private Sub DisplayFlaggedPicture()
    Dim tBM As BITMAP
        
    GetObjectAPI FlagPict.Handle, Len(tBM), tBM
    Debug.Print tBM.bmWidth
    picDesign.PaintPicture FlagPict, 11, 0, tBM.bmWidth, tBM.bmHeight
End Sub

Private Sub DisplayTimer(sTime As String, FontColor As Long, ShadowColor As Long)
    Dim tSZ As SIZE, tRECT As RECT, lhDC As Long
    
    picTime.Cls
    
    GetClientRect picTime.hwnd, tRECT
    GetTextExtentPoint32 picTime.hdc, sTime, Len(sTime), tSZ
    
    Dim xmid As Integer, ymid As Integer
    xmid = (tRECT.Right - tSZ.cx) / 2
    ymid = (tRECT.Bottom - tSZ.cy) / 2
    SetTextColor picTime.hdc, FontColor
    TextOut picTime.hdc, xmid + 1, ymid + 1, sTime, Len(sTime)
    
    If ShadowColor <> -1 Then ' if shadow is equal to -1 then no shadow
        SetTextColor picTime.hdc, ShadowColor
        TextOut picTime.hdc, xmid, ymid, sTime, Len(sTime)
    End If
End Sub

Private Sub AnswerCompleted(text As String, FontColor As Long, ShadowColor As Long)
    Dim tSZ As SIZE, tRECT As RECT, lhDC As Long
    
    picCompleted.Cls
    
    GetClientRect picCompleted.hwnd, tRECT
    GetTextExtentPoint32 picCompleted.hdc, text, Len(text), tSZ
    
    Dim xmid As Integer, ymid As Integer
    xmid = (tRECT.Right - tSZ.cx) / 2
    ymid = (tRECT.Bottom - tSZ.cy) / 2
    SetTextColor picCompleted.hdc, FontColor
    TextOut picCompleted.hdc, xmid + 1, ymid + 1, text, Len(text)
    
    If ShadowColor <> -1 Then ' if shadow is equal to -1 then no shadow
        SetTextColor picCompleted.hdc, ShadowColor
        TextOut picCompleted.hdc, xmid, ymid, text, Len(text)
    End If
End Sub

Private Sub tmrTimer_Timer()
    If Second(Time) <> ss Then
        ss = ss + 1
        If ss = 60 Then
            mm = mm + 1
            ss = 0
            TotalMins = TotalMins + 1
            If mm = 60 Then
                hh = hh + 1
                mm = 0
                If hh = 24 Then
                    hh = 0
                End If
            End If
        End If
    End If
    
    pgbSecond.Value = ss
    pgbMinute.Value = TotalMins
    DisplayTimer "Time : " & Format$(hh, "00:") _
        & Format$(mm, "00:") & Format$(ss, "00"), vbBlack, -1
    
    If TotalMins = MyExam.MaxTime Then
        MsgBox "Times up!!!", vbOKOnly Or vbInformation, "Examination"
        MyExam.ComputeScore
        StartCategory
    End If
End Sub

Private Sub StartCategory()
    Dim Mess As String, cat As String
        
    If MyExam.CurCategory <> -1 Then
        tmrTimer.Enabled = False
        Reset
        cat = AvailCategory(MyExam.CurCategory).Description
        MyExam.NextCategory
        Mess = "You are given " & MyExam.MaxTime _
            & " minute(s) to answer all the questions."
        MsgBox Mess, vbOKOnly Or vbInformation, _
            "Category : " & cat
        picDesign.Cls
        cat = "Category : " & cat
        sbStatusBar.Panels(1).text = cat
        pgbMinute.Max = MyExam.MaxTime
        AnswerCompleted "Completed at - 0/" & MyExam.TotalItems, vbBlack, -1
        cmdPos_Click 0
        tmrTimer.Enabled = True
    Else
        tmrTimer.Enabled = False
        cmdNextCategory.Enabled = False
        frmScore.Show vbModal
    End If
End Sub

Private Sub Reset()
    CurPos = 0
    ss = 0
    mm = 0
    hh = 0
    Set MyAnswer = Nothing
    Set DataInfo = Nothing
    pgbSecond.Value = 0
    pgbMinute.Value = 0
    TotalMins = 0
    TotalAns = 0
    
    Dim i As Integer
    For i = 0 To cbOption.Count - 1
        cbOption(i).Press = False
        lblOption(i).BackStyle = 0
    Next i
End Sub
