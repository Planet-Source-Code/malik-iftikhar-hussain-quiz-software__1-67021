VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExamDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Primary.mdb
'   - all correct answers are stored in Option 1.
'   - you can add as many categories as you like.

Public Title      As String
Public Direction  As String
Public TotalItems As Integer
Public MaxTime    As Integer

Public CurCategory As Integer

Dim DB As Database
Dim RS As Recordset

Public Sub InitExam(Filename As String)
    On Error GoTo ErrOpen
    
    Set DB = OpenDatabase(Filename)
    
    Dim tdfLoop As TableDef, oTemp As Object
    
    ' read all tables
    For Each tdfLoop In DB.TableDefs
        If Trim(UCase(Left(tdfLoop.Name, 10))) = "CATEGORY@1" Then
            Set oTemp = New cCategory
            oTemp.Table = tdfLoop.Name
            oTemp.Description = Mid$(tdfLoop.Name, 12, Len(tdfLoop.Name))
            oTemp.Items = 0
            oTemp.Score = 0
            AvailCategory.Add oTemp
        End If
    Next tdfLoop
    
    Set RS = DB.OpenRecordset("select * from Title")
    Title = RS![ExamTitle]
    CurCategory = 1
    Exit Sub
ErrOpen:
    MsgBox Err.Description, vbCritical Or vbOKOnly, "Error"
    End
End Sub

Public Sub Category(s As String)
    On Error GoTo ErrCategory
    
    Static NextItem As Integer
    Dim MaxRecord As Integer, temp As String
    
    temp = Replace(s, "Category@1", "Category@2")
    temp = "select * from [" & temp & "];"
    Set RS = DB.OpenRecordset(temp)
    Direction = Space(5) & RS![Direction]
    TotalItems = RS![Items]
    MaxTime = RS![Time]
    NextItem = NextItem + 1
    AvailCategory(NextItem).Items = TotalItems
    s = "select * from [" & s & "];"
    Set RS = DB.OpenRecordset(s)
    
    If RS.RecordCount = 0 Then
        MsgBox "No record found!", vbInformation Or vbOKOnly, _
            "Category : " & AvailCategory(CurCategory).Description
        End
    Else
        Dim ArrTempQ() As Integer   ' Array Questions
        Dim ArrTempO() As Integer   ' Array Options
        
        Dim TotalRecords As Long
        Dim oTemp As Object, i As Integer, j As Integer
        
        RS.MoveFirst
        RS.MoveLast
        TotalRecords = RS.AbsolutePosition + 1
        
        ReDim ArrTempQ(TotalRecords) As Integer
        ReDim ArrTempO(4) As Integer
        
        Shuffle ArrTempQ()
        frmMain.imcListAns.ComboItems.Clear
        For i = LBound(ArrTempQ()) To UBound(ArrTempQ()) - 1
            Set oTemp = New cData
            oTemp.ItemID = ArrTempQ(i)
            Shuffle ArrTempO()
            For j = LBound(ArrTempO()) To UBound(ArrTempO()) - 1
                oTemp.OptionID.Add ArrTempO(j)
                ArrTempO(j) = 0
            Next j
            DataInfo.Add oTemp
        Next i
        
        For i = LBound(ArrTempQ()) To TotalItems - 1
            frmMain.imcListAns.ComboItems.Add , , Format$(i + 1, "000"), "imgLightOff"
            frmMain.imcListAns.ComboItems(1).Selected = True
            Set oTemp = New cAnswer
            oTemp.Answer = 0
            oTemp.Selected = 0
            oTemp.Flag = False
            MyAnswer.Add oTemp
        Next i
    End If
    Exit Sub
ErrCategory:
    MsgBox Err.Description, vbCritical Or vbOKOnly, "Error"
End Sub

Public Sub Question(ItemID As Integer)
    RS.MoveLast
    RS.MoveFirst
    RS.FindFirst "[ItemID]=" & ItemID
    
    If RS.NoMatch Then
        RS.MoveFirst
        Question RS.AbsolutePosition
    Else
        With frmMain
            Dim sTemp As String

            .lblQuestion.Caption = RS.Fields("Question").Value
            sTemp = "[Option " & DataInfo(ItemID).OptionID(1) & "]"
            .lblOption(0).Caption = RS.Fields(sTemp).Value
            sTemp = "[Option " & DataInfo(ItemID).OptionID(2) & "]"
            .lblOption(1).Caption = RS.Fields(sTemp).Value
            sTemp = "[Option " & DataInfo(ItemID).OptionID(3) & "]"
            .lblOption(2).Caption = RS.Fields(sTemp).Value
            sTemp = "[Option " & DataInfo(ItemID).OptionID(4) & "]"
            .lblOption(3).Caption = RS.Fields(sTemp).Value
        End With
    End If
End Sub

Public Sub ComputeScore()
    Dim n As Integer
    
    n = IIf(CurCategory = -1, AvailCategory.Count + 1, CurCategory)
    
    Dim i As Integer, pos As Integer, Score As Integer
    
    Score = 0
    For i = 1 To MyAnswer.Count
        ' if answer is equal to 1 then your answer is correct
        If MyAnswer(i).Answer = 1 Then
            Score = Score + 1
        End If
    Next i
    
    AvailCategory(n - 1).Score = Score
End Sub

Public Sub NextCategory()
    If CurCategory <= AvailCategory.Count Then
        Category AvailCategory(CurCategory).Table
    End If
    
    CurCategory = CurCategory + 1
    If CurCategory > AvailCategory.Count Then
        CurCategory = -1 ' no more category
    End If
End Sub

Public Sub CleanUp()
    RS.Close
    DB.Close
End Sub
