VERSION 5.00
Begin VB.Form frmDirection 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FAFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3720
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Direction"
   Begin VB.Label lblExit 
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
      Height          =   315
      Left            =   1260
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3900
      Width           =   1335
   End
End
Attribute VB_Name = "frmDirection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim lhDC As Long
    Dim tRECT As RECT, tBM As BITMAP
    Dim i As Integer, step As Integer
    Dim Temp As Picture
    On Error Resume Next
    
    Set Temp = LoadPicture(ImagePath & "notebook.bmp")
    GetClientRect Me.hwnd, tRECT
    GetObjectAPI Temp, Len(tBM), tBM
    step = 0
    For i = 0 To tRECT.Right / tBM.bmWidth
        Me.PaintPicture Temp, step, 0, tBM.bmWidth, tBM.bmHeight
        step = step + tBM.bmWidth
    Next i
    If Not (lhDC = 0) Then DeleteDC lhDC
    
    Dim PencilTailWidth As Long, PencilHeadWidth As Long
    Dim OffsetPencilX As Long, OffsetPencilY As Long
    Dim Sprite As Picture, Mask As Picture
    
    Set Sprite = LoadPicture(ImagePath & "pheadsprite.bmp")
    Set Mask = LoadPicture(ImagePath & "pheadmask.bmp")
    GetObjectAPI Sprite.Handle, Len(tBM), tBM
    OffsetPencilX = 10
    OffsetPencilY = tBM.bmHeight + 5
    PencilHeadWidth = tBM.bmWidth
    Me.PaintPicture Mask, tRECT.Right - tBM.bmWidth - OffsetPencilX, OffsetPencilY, tBM.bmWidth, tBM.bmHeight, , , , , vbSrcAnd
    Me.PaintPicture Sprite, tRECT.Right - tBM.bmWidth - OffsetPencilX, OffsetPencilY, tBM.bmWidth, tBM.bmHeight, , , , , vbSrcInvert
    
    Set Sprite = LoadPicture(ImagePath & "ptailsprite.bmp")
    Set Mask = LoadPicture(ImagePath & "ptailmask.bmp")
    GetObjectAPI Sprite.Handle, Len(tBM), tBM
    PencilTailWidth = tBM.bmWidth
    ' create transparent bitmap
    Me.PaintPicture Mask, OffsetPencilX, OffsetPencilY, tBM.bmWidth, tBM.bmHeight, , , , , vbSrcAnd
    Me.PaintPicture Sprite, OffsetPencilX, OffsetPencilY, tBM.bmWidth, tBM.bmHeight, , , , , vbSrcInvert
    
    Set Temp = LoadPicture(ImagePath & "pbody.bmp")
    GetObjectAPI Temp.Handle, Len(tBM), tBM
    step = 0
    For i = 0 To (tRECT.Right - PencilHeadWidth - PencilTailWidth - OffsetPencilX * 2) / tBM.bmWidth
        Me.PaintPicture Temp, step + OffsetPencilX + PencilTailWidth, OffsetPencilY, tBM.bmWidth, tBM.bmHeight, , , , , vbSrcCopy
        step = step + tBM.bmWidth
    Next i
    
    Dim tSize As SIZE, OldTextColor As Long
    Dim hFont As Long, tLf As LOGFONT, OldFont As Long
    
    tLf.lfWeight = &H2BC ' Bold
    tLf.lfWidth = &HE
    tLf.lfHeight = &H1A
    hFont = CreateFontIndirect(tLf)
    OldFont = SelectObject(Me.hdc, hFont)
    GetTextExtentPoint32 Me.hdc, Me.Tag, Len(Me.Tag), tSize
    OldTextColor = SetTextColor(Me.hdc, &HC0C0FF)
    TextOut Me.hdc, (tRECT.Right - tSize.cx) / 2, OffsetPencilY + 2, Me.Tag, Len(Me.Tag)
    SetTextColor Me.hdc, OldTextColor
    SelectObject Me.hdc, OldFont
    If Not (hFont = 0) Then DeleteObject hFont
    Rectangle Me.hdc, 0, 0, tRECT.Right, tRECT.Bottom
    lblExit.Left = (tRECT.Right - lblExit.Width) / 2
    lblExit.Top = tRECT.Bottom - lblExit.Height - 5
    Set lblExit.MouseIcon = LoadResPicture(101, vbResCursor)
    
    tRECT.Left = tRECT.Left + 10
    tRECT.Top = tRECT.Top + 85
    tRECT.Right = tRECT.Right - 10
    tRECT.Bottom = tRECT.Bottom - 35
    
    DrawText Me.hdc, MyExam.Direction, Len(MyExam.Direction), _
        tRECT, DT_LEFT Or DT_WORDBREAK
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

