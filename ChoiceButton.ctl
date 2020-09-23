VERSION 5.00
Begin VB.UserControl ChoiceButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   MaskColor       =   &H00000000&
   MousePointer    =   99  'Custom
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   25
End
Attribute VB_Name = "ChoiceButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_KEYDOWN As Long = &H100

Private Type BITMAP         ' The BITMAP structure defines the type, width, height, color format, and bit values of a bitmap.
    bmType As Long          ' Specifies the bitmap type. This member must be zero.
    bmWidth As Long         ' Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
    bmHeight As Long        ' Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
    bmWidthBytes As Long    ' Specifies the number of bytes in each scan line. This value must be divisible by 2,
                            '   because the system assumes that the bit values of a bitmap form an array that is word aligned.
    bmPlanes As Integer     ' Specifies the count of color planes.
    bmBitsPixel As Integer  ' Specifies the number of bits required to indicate the color of a pixel.
    bmBits As Long          ' Pointer to the location of the bit values for the bitmap. The bmBits member must be a long pointer
                            '   to an array of character (1-byte) values.
End Type

'  The GetObject function retrieves information for the specified graphics object.
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'  The PostMessage function places (posts) a message in the message queue associated with the thread that created the specified window and then returns without waiting for the thread to process the message.
' Messages in a message queue are retrieved by calls to the GetMessage or PeekMessage function.
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const resOffsetButtonUp = 101           ' Button Up
Private Const resOffsetButtonDn = 105           ' Button Down
Private Const resOffsetButtonUpFocus = 109      ' Button Up Focus
Private Const resOffsetButtonDnFocus = 113      ' Button Down Focus
Private Const resMaskButtonUp = 120
Private Const resMaskButtonDn = 121

Public Enum SelectionConstant
    cb_A    ' Button A
    cb_B    ' Button B
    cb_C    ' Button C
    cb_D    ' Button D
    cb_E
End Enum

Private Type cbProperties
    cbSelection As SelectionConstant
    cbPress As Boolean
End Type

Dim IsFocus As Boolean
Dim MyProp As cbProperties
Dim bW As Integer, bH As Integer

Event Click()

Public Property Get Selection() As SelectionConstant
    Selection = MyProp.cbSelection
End Property

Public Property Let Selection(cbSelection As SelectionConstant)
    If cbSelection <> MyProp.cbSelection Then
        MyProp.cbSelection = cbSelection
        PropertyChanged "Selection"
        RedrawButton
    End If
End Property

Public Property Get Press() As Boolean
    Press = MyProp.cbPress
End Property

Public Property Let Press(cbPress As Boolean)
    If cbPress <> MyProp.cbPress Then
        MyProp.cbPress = cbPress
        PropertyChanged "Press"
        RedrawButton
    End If
End Property

Private Sub UserControl_Click()
    Press = Not Press
    RaiseEvent Click
End Sub

Private Sub UserControl_ExitFocus()
    Dim TempValue As Integer
    
    If IsFocus Then
        IsFocus = False
        TempValue = IIf(Press, resOffsetButtonDn, resOffsetButtonUp)
        Set UserControl.Picture = LoadResPicture(TempValue + Selection, vbResBitmap)
    End If
End Sub

Private Sub UserControl_GotFocus()
    Dim TempValue As Integer
    
    If Not IsFocus Then
        IsFocus = True
        TempValue = IIf(Press, resOffsetButtonDnFocus, resOffsetButtonUpFocus)
        Set UserControl.Picture = LoadResPicture(TempValue + Selection, vbResBitmap)
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cHwnd As Long
    
    cHwnd = UserControl.ContainerHwnd
    
    Select Case KeyCode
    Case Is = vbKeyRight
        KeyCode = 0
        PostMessage cHwnd, WM_KEYDOWN, ByVal &H27, ByVal &H4D0001
    Case Is = vbKeyDown
        KeyCode = 0
        PostMessage cHwnd, WM_KEYDOWN, ByVal &H28, ByVal &H500001
    Case Is = vbKeyLeft
        KeyCode = 0
        PostMessage cHwnd, WM_KEYDOWN, ByVal &H25, ByVal &H4B0001
    Case Is = vbKeyUp
        KeyCode = 0
        PostMessage cHwnd, WM_KEYDOWN, ByVal &H26, ByVal &H480001
    Case vbKeySpace, vbKeyReturn
        UserControl_Click
    End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyProp.cbSelection = PropBag.ReadProperty("Selection", cb_A)
    MyProp.cbPress = PropBag.ReadProperty("Press", False)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = bW * Screen.TwipsPerPixelX
    UserControl.Height = bH * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Show()
    Dim tBM As BITMAP
    
    RedrawButton
    GetObjectAPI UserControl.Picture.Handle, Len(tBM), tBM
    bW = tBM.bmWidth: bH = tBM.bmHeight
    UserControl.Width = bW * Screen.TwipsPerPixelX
    UserControl.Height = bH * Screen.TwipsPerPixelY
    Set UserControl.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Selection", MyProp.cbSelection, cb_A)
    Call PropBag.WriteProperty("Press", MyProp.cbPress, False)
End Sub

Private Sub RedrawButton()
    Dim TempValue As Integer
        
    If Not IsFocus Then
        TempValue = IIf(Press, resOffsetButtonDn, resOffsetButtonUp)
    Else
        TempValue = IIf(Press, resOffsetButtonDnFocus, resOffsetButtonUpFocus)
    End If
    
    Set UserControl.Picture = LoadResPicture(TempValue + Selection, vbResBitmap)
    TempValue = IIf(Press, resMaskButtonDn, resMaskButtonUp)
    Set UserControl.MaskPicture = LoadResPicture(TempValue, vbResBitmap)
End Sub
