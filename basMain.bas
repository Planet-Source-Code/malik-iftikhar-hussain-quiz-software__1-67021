Attribute VB_Name = "basMain"
Option Explicit

Public ImagePath As String

Public MyExam         As New cExamDB
Public MyAnswer       As New Collection
Public AvailCategory  As New Collection
Public DataInfo As New Collection

Public Sub Main()
    MyExam.InitExam App.Path & "\primary.mdb"
    ImagePath = App.Path & "\Image\"
    Load frmMain
    frmMain.Show
End Sub

Public Sub Shuffle(ByRef Data() As Integer)
    Dim i As Integer, IsInArray As Boolean
    Dim MaxValue As Integer, TempValue, step As Integer

    step = 0
    Randomize
    
    MaxValue = UBound(Data())
    Do While step < MaxValue
        IsInArray = False
        TempValue = Int((MaxValue * Rnd) + 1)
        For i = 0 To step
            If Data(i) = TempValue Then
                IsInArray = True
                Exit For
            End If
        Next i
        If Not IsInArray Then
            Data(step) = TempValue
            step = step + 1
        End If
    Loop
End Sub

Public Function BitmapToPicture(ByVal hBMP As Long) As IPicture

   If (hBMP = 0) Then Exit Function

   Dim NewPic As Picture, tPicConv As PictDesc, IGuid As Guid

   ' Fill PictDesc structure with necessary parts:
   With tPicConv
      .cbSizeofStruct = Len(tPicConv)
      .picType = vbPicTypeBitmap
      .hImage = hBMP
   End With

   ' Fill in IDispatch Interface ID
   With IGuid
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Create a picture object:
   OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
   
   ' Return it:
   Set BitmapToPicture = NewPic
End Function

Public Sub DrawLine(lhDC As Long, X1 As Integer, Y1 As Integer, _
                     X2 As Integer, Y2 As Integer, nWidth As Integer, cColor As Long)

    Dim hPen As Long, OldPen As Long, tPT As POINTAPI
    
    hPen = CreatePen(PS_SOLID, nWidth, cColor)
    OldPen = SelectObject(lhDC, hPen)
    MoveToEx lhDC, X1, Y1, tPT
    LineTo lhDC, X2, Y2
    SelectObject lhDC, OldPen
    DeleteObject hPen
End Sub

