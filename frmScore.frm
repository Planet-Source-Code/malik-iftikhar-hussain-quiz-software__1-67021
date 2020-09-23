VERSION 5.00
Begin VB.Form frmScore 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6465
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<< Exit >>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2700
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3360
      Width           =   1200
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim tSZ As Size, tRECT As RECT
    Dim i As Integer, step As Integer
    
    GetClientRect Me.hWnd, tRECT
    GetTextExtentPoint32 Me.hdc, "H", Len("H"), tSZ
    step = (tRECT.Bottom - tRECT.Top) / tSZ.cy
    lblExit.Left = (tRECT.Right - lblExit.Width) / 2
    lblExit.Top = tSZ.cy * (step - 1)
    Set lblExit.MouseIcon = LoadResPicture(101, vbResCursor)
    DrawLine Me.hdc, CInt(tRECT.Left), tSZ.cy * 2, CInt(tRECT.Right), _
            tSZ.cy * 2, 2, vbRed
    For i = 3 To step
        DrawLine Me.hdc, CInt(tRECT.Left), tSZ.cy * i, CInt(tRECT.Right), _
            tSZ.cy * i, 1, &HFF0000
    Next i
    
    PutText Me.hdc, MyExam.Title, 5, CInt(tSZ.cy), &H2BC, vbBlue
    PutText Me.hdc, "Category", 10, 3 * tSZ.cy, &H2BC, vbRed
    PutText Me.hdc, "Items", tRECT.Right - (Len("Items") + 15) * tSZ.cx, 3 * tSZ.cy, &H2BC, vbRed
    PutText Me.hdc, "Score", tRECT.Right - (Len("Score") + 2) * tSZ.cx, 3 * tSZ.cy, &H2BC, vbRed
    
    For i = 1 To AvailCategory.Count
        SetTextAlign Me.hdc, TA_LEFT
        PutText Me.hdc, " - " & AvailCategory(i).Description, 20, (i + 4) * tSZ.cy, &H190, vbBlack
        SetTextAlign Me.hdc, TA_RIGHT
        PutText Me.hdc, CStr(AvailCategory(i).Score), tRECT.Right - 4 * tSZ.cx + 8, (i + 4) * tSZ.cy, &H190, vbBlack
        PutText Me.hdc, CStr(AvailCategory(i).Items), tRECT.Right - 4 * tSZ.cx - 108, (i + 4) * tSZ.cy, &H190, vbBlack
    Next i
End Sub

Private Sub PutText(lhDC As Long, sText As String, x As Integer, y As Integer, _
                    nWeight As Integer, color As Long)
                    
    Dim hFont As Long, OldFont As Long, tLf As LOGFONT

    tLf.lfWeight = nWeight
    tLf.lfQuality = PROOF_QUALITY
    hFont = CreateFontIndirect(tLf)
    OldFont = SelectObject(Me.hdc, hFont)
    SetTextColor Me.hdc, color
    TextOut lhDC, x, y, sText, Len(sText)
    SelectObject Me.hdc, OldFont
    DeleteObject hFont
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

