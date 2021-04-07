VERSION 5.00
Begin VB.UserControl axComboBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "axComboBox.ctx":0000
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   45
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2085
      Begin VB.ListBox Lst 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   225
         Sorted          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1635
      End
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   45
      ScaleHeight     =   600
      ScaleWidth      =   2085
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   2085
      Begin VB.PictureBox picButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   150
         Width           =   285
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   260
         Left            =   195
         TabIndex        =   0
         Top             =   150
         Width           =   1230
      End
   End
End
Attribute VB_Name = "axComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------TYPES---------------
Private Type POINTAPI
    x As Long
    y As Long
End Type
    
Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
    
Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type
    
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'-------ENUMS---------------
Private Enum Blends
    RGBBlend = 0
    HSLBlend = 1
End Enum

Public Enum iSelectMode
    SingleClick = 0
    DoubleClick = 1
End Enum

Public Enum eEnterKeyBehavior
    eNone = 0
    eKeyTab = 1
    eAddItem = 2
End Enum


'-----DECLARACIONES----------------------------------------------------------
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
'Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long

' recupera el estilo del Listbox
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' cambia el estilo
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' refresca y vuelve a redibujar el control
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)
'--------------------------------------------------------

'------CONTROL--------------------------
Private WithEvents tmrMouseMove As Timer
Attribute tmrMouseMove.VB_VarHelpID = -1

'------VARIABLES-----------------
Dim OnLoad          As Boolean
Dim m_BackColor     As OLE_COLOR
Dim m_BorderColor   As OLE_COLOR
Dim m_FocusColor    As OLE_COLOR
Dim m_SelTextFocus  As Boolean
Dim m_ButtonColor   As OLE_COLOR
Dim Expanded        As Boolean
Dim m_Elements      As Integer
Dim m_SelectMode    As iSelectMode
Dim m_KeyBehavior   As eEnterKeyBehavior

'You have to have MSScripting Runtime referenced : WshShell.SendKeys "{Tab}"
Dim WshShell        As Object

'------Private Variables---------
Private lBottomR As Long
Private lBottomG As Long
Private lBottomB As Long
Private lTopR As Long
Private lTopG As Long
Private lTopB As Long
Private Col1  As Long
Private Col2  As Long
Private LstH  As Integer

'------CONSTANTES-----------------
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&
Private Const m_def_BackColor = &HFFFFFF
Private Const m_def_BorderColor = &H808080
Private Const m_def_FocusColor = &HFFFFC0
Private Const m_def_ButtonColor = &H404040
Private Const m_def_Items = 10
' constantes para SetWindowPos
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
' para GetWindowLong - SetWindowLong
Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000

Private Const LB_ADDSTRING As Long = &H180

'------EVENTOS-----------------
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
'Public Event KeyDown(KeyCode As Integer, Shift As Integer)
'Public Event KeyPress(KeyAscii As Integer)
Public Event EnterKeyPress()



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Agrega un elemento a un control Listbox o ComboBox, o una fila a un control Grid."
If IsMissing(Index) Then
    Lst.AddItem Item
Else
    Lst.AddItem Item, Index
End If
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Quita un elemento de un control ListBox o ComboBox, o una fila de un control Grid."
On Error Resume Next
    Lst.RemoveItem Index
End Sub

Public Sub Clear()
    Lst.Clear

End Sub


'------Helper Functions----------------
Private Function CreateGradient(Width As Long, Height As Long, LeftToRight As Boolean, LeftTopColor As Long, RightBottomColor As Long, BlendType As Blends) As StdPicture
    Dim hBmp As Long, Bits() As Byte
    Dim RS As Byte, GS As Byte, BS As Byte
    Dim RE As Byte, GE As Byte, BE As Byte
    Dim HS As Single, SS As Single, lS As Single
    Dim HE As Single, SE As Single, LE As Single
    Dim rc As Byte, GC As Byte, BC As Byte
    Dim x As Long, y As Long
    ReDim Bits(0 To 3, 0 To Width - 1, 0 To Height - 1)
    
    RgbCol LeftTopColor, RS, GS, BS
    RgbCol RightBottomColor, RE, GE, BE
    
    If BlendType = RGBBlend Then
        If LeftToRight Then
            For x = 0 To Width - 1
                rc = (1& * RS - RE) * ((Width - 1 - x) / (Width - 1)) + RE
                GC = (1& * GS - GE) * ((Width - 1 - x) / (Width - 1)) + GE
                BC = (1& * BS - BE) * ((Width - 1 - x) / (Width - 1)) + BE
                For y = 0 To Height - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        Else
            For y = 0 To Height - 1
                rc = (1& * RS - RE) * ((Height - 1 - y) / (Height - 1)) + RE
                GC = (1& * GS - GE) * ((Height - 1 - y) / (Height - 1)) + GE
                BC = (1& * BS - BE) * ((Height - 1 - y) / (Height - 1)) + BE
                For x = 0 To Width - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        End If
    ElseIf BlendType = HSLBlend Then
        RGBToHSL RS, GS, BS, HS, SS, lS
        RGBToHSL RE, GE, BE, HE, SE, LE
        If LeftToRight Then
            For x = 0 To Width - 1
                HSLToRGB (1& * HS - HE) * ((Width - 1 - x) / (Width - 1)) + HE, _
                        (1& * SS - SE) * ((Width - 1 - x) / (Width - 1)) + SE, _
                        (1& * lS - LE) * ((Width - 1 - x) / (Width - 1)) + LE, _
                        rc, GC, BC
                For y = 0 To Height - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        Else
            For y = 0 To Height - 1
                HSLToRGB (1& * HS - HE) * ((Height - 1 - y) / (Height - 1)) + HE, _
                        (1& * SS - SE) * ((Height - 1 - y) / (Height - 1)) + SE, _
                        (1& * lS - LE) * ((Height - 1 - y) / (Height - 1)) + LE, _
                        rc, GC, BC
                For x = 0 To Width - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        End If
    End If

    Dim BI As BITMAPINFO
    With BI.bmiHeader
        .biSize = Len(BI.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    hBmp = CreateBitmap(Width, Height, 1&, 32&, ByVal 0)
    SetDIBits 0&, hBmp, 0, Abs(BI.bmiHeader.biHeight), Bits(0, 0, 0), BI, DIB_RGB_COLORS

    Dim IGuid As Guid, PicDst As PictDesc
    
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    OleCreatePictureIndirect PicDst, IGuid, True, CreateGradient
End Function

Private Sub DrawBorders()
UserControl.Cls
Dim Rgn As Long

With UserControl
    Rgn = CreateRoundRectRgn(0, 0, .Width, .Height, 0, 0)
    SetWindowRgn .hwnd, Rgn, True
    DeleteObject Rgn
    .DrawWidth = 1
    .ForeColor = m_BorderColor
    RoundRect .hDC, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0
End With

With picList
    Rgn = CreateRoundRectRgn(0, 0, .Width, .Height, 0, 0)
    SetWindowRgn .hwnd, Rgn, True
    DeleteObject Rgn
    .DrawWidth = 1
    .ForeColor = m_BorderColor
    RoundRect .hDC, 0, 0, .ScaleWidth, .ScaleHeight, 0, 0
End With

Dim lng_Estilo As Long

    With Lst
        '.Appearance = 0 ' flat
        lng_Estilo = GetWindowLong(.hwnd, GWL_STYLE)
        lng_Estilo = lng_Estilo And Not WS_BORDER ' sin borde
        ' aplica
        SetWindowLong .hwnd, GWL_STYLE, lng_Estilo
        ' refresh
        SetWindowPos .hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or _
                                           SWP_NOACTIVATE Or _
                                           SWP_NOMOVE Or _
                                           SWP_NOOWNERZORDER Or _
                                           SWP_NOSIZE Or _
                                           SWP_NOZORDER
    End With

End Sub

Private Sub DrawPin(isMouseOver As Boolean)
  Dim i As Integer
  Dim iHorizontal1 As Integer
  Dim iHorizontal2 As Integer
  Dim iVertical As Integer
  Dim cColor As OLE_COLOR
  
  picButton.Cls
  
  If Expanded = False Then
        Col1 = m_ButtonColor
        Col2 = &HFFFFFF
  Else
        Col1 = &HFFFFFF
        Col2 = m_ButtonColor
  End If
  
  lBottomR = (Col1 And &HFF&)
  lBottomG = (Col1 And &HFF00&) / &H100
  lBottomB = (Col1 And &HFF0000) / &H10000
  
  lTopR = (Col2 And &HFF&)
  lTopG = (Col2 And &HFF00&) / &H100
  lTopB = (Col2 And &HFF0000) / &H10000

Set picButton.Picture = CreateGradient(picButton.Width / Screen.TwipsPerPixelX, picButton.Height / Screen.TwipsPerPixelY, False, RGB(lTopR, lTopG, lTopB), RGB(lBottomR, lBottomG, lBottomB), RGBBlend)

  If isMouseOver = False Then
      cColor = m_BackColor
  Else
      cColor = m_FocusColor
  End If
  

If Expanded = False Then
      iHorizontal1 = 145 '210
      iHorizontal2 = 130 '195
      iVertical = 45
      For i = 1 To 2
          ' 1st Line of 1st Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 45), iVertical)-(picButton.Width - (iHorizontal1 + 15), iVertical), cColor
          picButton.Line (picButton.Width - (iHorizontal2 - 15), iVertical)-(picButton.Width - (iHorizontal2 - 45), iVertical), cColor
          iVertical = iVertical + 15
      
          ' 2nd Line of 1st Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 30), iVertical)-(picButton.Width - iHorizontal1, iVertical), cColor
          picButton.Line (picButton.Width - iHorizontal2, iVertical)-(picButton.Width - (iHorizontal2 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 1st Line of 2nd Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 15), iVertical)-(picButton.Width - (iHorizontal1 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 2nd Line of 2nd Arrow
          picButton.Line (picButton.Width - iHorizontal1, iVertical)-(picButton.Width - (iHorizontal1 - 15), iVertical), cColor
          iVertical = iVertical + 15
      Next

Else
      iHorizontal1 = 145
      iHorizontal2 = 130
      iVertical = 45
      For i = 1 To 2
          ' 1st Line of 1st Arrow
          picButton.Line (picButton.Width - iHorizontal1, iVertical)-(picButton.Width - (iHorizontal1 - 15), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 2nd Line of 1st Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 15), iVertical)-(picButton.Width - (iHorizontal1 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 1st Line of 2nd Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 30), iVertical)-(picButton.Width - iHorizontal1, iVertical), cColor
          picButton.Line (picButton.Width - iHorizontal2, iVertical)-(picButton.Width - (iHorizontal2 - 30), iVertical), cColor
          iVertical = iVertical + 15
          
          ' 2nd Line of 2nd Arrow
          picButton.Line (picButton.Width - (iHorizontal1 + 45), iVertical)-(picButton.Width - (iHorizontal1 + 15), iVertical), cColor
          picButton.Line (picButton.Width - (iHorizontal2 - 15), iVertical)-(picButton.Width - (iHorizontal2 - 45), iVertical), cColor
          iVertical = iVertical + 15
      Next
End If

End Sub

Private Sub FindPointer()
    Dim pt As POINTAPI

    GetCursorPos pt
    
    If WindowFromPointXY(pt.x, pt.y) <> picButton.hwnd Then
        Call DrawPin(False)
        tmrMouseMove.Enabled = False
    Else
       Call DrawPin(True)
    End If
    
    Call DrawBorders
End Sub

Private Sub HSLToRGB(ByVal H As Single, ByVal S As Single, ByVal l As Single, R As Byte, g As Byte, b As Byte)
    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single
    
    If S = 0 Then
        rR = l: rG = l: rB = l
    Else
        If l <= 0.5 Then
            Min = l * (1 - S)
        Else
            Min = l - S * (1 - l)
        End If
        Max = 2 * l - Min
       
        If (H < 1) Then
            rR = Max
            If (H < 0) Then
                rG = Min
                rB = rG - H * (Max - Min)
            Else
                rB = Min
                rG = H * (Max - Min) + rB
            End If
        ElseIf (H < 3) Then
            rG = Max
            If (H < 2) Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
            Else
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (H < 4) Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
            Else
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
    End If
    R = rR * 255: g = rG * 255: b = rB * 255
End Sub

Private Sub RgbCol(Col As Long, ByRef R As Byte, ByRef g As Byte, ByRef b As Byte)
    R = Col And &HFF&
    g = (Col And &HFF00&) \ &H100&
    b = (Col And &HFF0000) \ &H10000
End Sub

Private Sub RGBToHSL(ByVal R As Byte, ByVal g As Byte, ByVal b As Byte, H As Single, S As Single, l As Single)
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single

    rR = R / 255: rG = g / 255: rB = b / 255

    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2
    If Max = Min Then
        S = 0
        H = 0
    Else
        If l <= 0.5 Then
            S = (Max - Min) / (Max + Min)
        Else
            S = (Max - Min) / (2 - Max - Min)
        End If
        
        delta = Max - Min
        If rR = Max Then
            H = (rG - rB) / delta
        ElseIf rG = Max Then
            H = 2 + (rB - rR) / delta
        ElseIf rB = Max Then
            H = 4 + (rR - rG) / delta
        End If
    End If
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR > rG) Then
        If (rR > rB) Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If (rB > rG) Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR < rG) Then
        If (rR < rB) Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If (rB < rG) Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If
End Function

Private Sub UserControlsCreate()
Set tmrMouseMove = UserControl.Controls.Add("VB.Timer", "tmrMouseMove")
tmrMouseMove.Interval = 100
tmrMouseMove.Enabled = False

End Sub

Private Sub Lst_DblClick()
If OnLoad = False Then
  If m_SelectMode = DoubleClick Then
      Txt.Text = Lst.Text
      picButton_Click
  Else
    Exit Sub
  End If
End If

RaiseEvent DblClick

End Sub

Private Sub Lst_Click()
If OnLoad = False Then
  If m_SelectMode = SingleClick Then
      Txt.Text = Lst.Text
      picButton_Click
  Else
    Exit Sub
  End If
End If

RaiseEvent Click
End Sub


Private Sub picButton_Click()
LstH = 225 + (195 * (m_Elements - 1))

OnLoad = False
Expanded = Not Expanded

With UserControl
    If Expanded = True Then
        picList.Visible = True
        picList.Move 0, Pic.Height + 1, .ScaleWidth - 2, (LstH + 10) / 15
        Lst.Move 2, 1, picList.ScaleWidth - 1, (LstH + 10) / 15
        .Height = LstH + 261 + 2
    Else
        .Height = 261
        picList.Visible = False
    End If
End With

Debug.Print UserControl.ScaleHeight & " / " & LstH & " / " & picList.Height

End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
tmrMouseMove.Enabled = True
End Sub

Private Sub tmrMouseMove_Timer()
Call FindPointer

End Sub

Private Sub Txt_Change()
RaiseEvent Change
End Sub

Private Sub Txt_GotFocus()
With Txt
  If m_SelTextFocus = True Then
      .SelStart = 0
      .SelLength = Len(.Text)
  End If
  .BackColor = m_FocusColor
End With

Pic.BackColor = m_FocusColor
End Sub

Private Sub Txt_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Select Case m_KeyBehavior
      Case Is = eNone
        'Do Nothing
      Case Is = eKeyTab
        WshShell.SendKeys "{Tab}"
        
      Case Is = eAddItem
        Dim iL As Integer
        With Lst
          iL = .ListCount
          .AddItem Txt.Text, iL
          'Abro List
          picButton_Click
        End With
  End Select

  RaiseEvent EnterKeyPress
End If
'---------------------------------
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Txt_LostFocus()
Txt.BackColor = m_BackColor
Pic.BackColor = m_BackColor

End Sub

Private Sub UserControl_Initialize()
Set WshShell = CreateObject("WScript.Shell")
Call UserControlsCreate
Call FindPointer

Expanded = False
OnLoad = True
Txt.BackColor = vbWhite
Lst.Clear
End Sub

Private Sub UserControl_InitProperties()
'Default Value Properties
m_BackColor = m_def_BackColor
m_BorderColor = m_def_BorderColor
m_FocusColor = m_def_FocusColor
m_ButtonColor = m_def_ButtonColor
m_SelectMode = SingleClick

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call FindPointer
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
With UserControl
    If Expanded = False Then
        .Height = 261
        Pic.Move 1, 1, .ScaleWidth - 2, .ScaleHeight - 2
        picButton.Move Pic.ScaleWidth - 255, 0, 250, 230
        Txt.Move 35, 10, Pic.ScaleWidth - 303, Pic.ScaleHeight - 2
    End If
End With

Call FindPointer
End Sub

''--------------------- PROPERTIES BAG ----------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
With PropBag
  Set Txt.Font = .ReadProperty("Font", Ambient.Font)
  Txt.ForeColor = .ReadProperty("ForeColor", vbButtonText)
  Txt.Enabled = .ReadProperty("Enabled", True)
  Txt.Text = .ReadProperty("Text", "")
  m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
  m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
  m_ButtonColor = .ReadProperty("ButtonColor", m_def_ButtonColor)
  m_FocusColor = .ReadProperty("BackColorOnFocus", m_def_FocusColor)
  m_SelTextFocus = .ReadProperty("SelTextOnFocus", False)
  m_Elements = .ReadProperty("ItemsInList", m_def_Items)
  Lst.List(Index) = .ReadProperty("List" & Index, "")
  Lst.ListIndex = .ReadProperty("ListIndex", 0)
  m_KeyBehavior = .ReadProperty("EnterKeyBehavior", eNone)
  m_SelectMode = .ReadProperty("ItemSelectMode", SingleClick)
End With

UserControl_Resize

  Txt.SelLength = PropBag.ReadProperty("SelLength", 0)
  Txt.SelStart = PropBag.ReadProperty("SelStart", 0)
  Txt.SelText = PropBag.ReadProperty("SelText", "")
End Sub

''--------------------- PROPERTIES BAG ----------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
With PropBag
   .WriteProperty "Font", Txt.Font, Ambient.Font
   .WriteProperty "ForeColor", Txt.ForeColor, vbButtonText
   .WriteProperty "Enabled", Txt.Enabled, True
   .WriteProperty "Text", Txt.Text, ""
   .WriteProperty "BackColor", m_BackColor, m_def_BackColor
   .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
   .WriteProperty "ButtonColor", m_ButtonColor, m_def_ButtonColor
   .WriteProperty "BackColorOnFocus", m_FocusColor, m_def_FocusColor
   .WriteProperty "SelTextOnFocus", m_SelTextFocus, False
   .WriteProperty "ItemsInList", m_Elements, m_def_Items
   .WriteProperty "ItemSelectMode", m_SelectMode, SingleClick
   .WriteProperty "List" & Index, Lst.List(Index), ""
   .WriteProperty "ListIndex", Lst.ListIndex, 0
   .WriteProperty "EnterKeyBehavior", m_KeyBehavior, eNone
   .WriteProperty "ItemSelectMode", m_SelectMode, SingleClick
End With

UserControl_Resize

  Call PropBag.WriteProperty("SelLength", Txt.SelLength, 0)
  Call PropBag.WriteProperty("SelStart", Txt.SelStart, 0)
  Call PropBag.WriteProperty("SelText", Txt.SelText, "")
End Sub

''------------------ PROPERTIES --------------------------
Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
m_BackColor = NewBackColor
Txt.BackColor = NewBackColor
UserControl.BackColor = NewBackColor
Pic.BackColor = NewBackColor
PropertyChanged "BackColor"
UserControl_Resize
End Property

Public Property Get BackColorOnFocus() As OLE_COLOR
BackColorOnFocus = m_FocusColor
End Property

Public Property Let BackColorOnFocus(ByVal NewColor As OLE_COLOR)
m_FocusColor = NewColor
PropertyChanged "BackColorOnFocus"
UserControl_Resize
End Property

Public Property Get BorderColor() As OLE_COLOR
BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
m_BorderColor = NewBorderColor
PropertyChanged "BorderColor"
UserControl_Resize
End Property

Public Property Get ButtonColor() As OLE_COLOR
ButtonColor = m_ButtonColor
End Property

Public Property Let ButtonColor(ByVal NewButtonColor As OLE_COLOR)
m_ButtonColor = NewButtonColor
PropertyChanged "ButtonColor"
UserControl_Resize
End Property

Public Property Get Enabled() As Boolean
Enabled = Txt.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Txt.Enabled = New_Enabled
PropertyChanged "Enabled"
End Property

Public Property Get EnterKeyBehavior() As eEnterKeyBehavior
EnterKeyBehavior = m_KeyBehavior
End Property

Public Property Let EnterKeyBehavior(ByVal NewBehavior As eEnterKeyBehavior)
m_KeyBehavior = NewBehavior
PropertyChanged "EnterKeyBehavior"
End Property

Public Property Get Font() As Font
Set Font = Txt.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set Txt.Font = New_Font
PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = Txt.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Txt.ForeColor = New_ForeColor
PropertyChanged "ForeColor"
End Property

Public Property Get ItemSelectMode() As iSelectMode
    ItemSelectMode = m_SelectMode
End Property

Public Property Let ItemSelectMode(ByVal new_Select As iSelectMode)
    m_SelectMode = new_Select
    PropertyChanged "ItemSelectMode"
End Property

Public Property Get ItemsInList() As Integer
ItemsInList = m_Elements
End Property

Public Property Let ItemsInList(ByVal vNewValue As Integer)
m_Elements = vNewValue
PropertyChanged "ItemsInList"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Devuelve el número de elementos en la parte de lista de un control."
    ListCount = Lst.ListCount
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Devuelve o establece los elementos contenidos en la parte de lista de un control."
    List = Lst.List(Index)
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Devuelve o establece el índice del elemento seleccionado actualmente en el control."
    ListIndex = Lst.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Lst.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Lst.List(Index) = New_List
    PropertyChanged "List"
End Property

Public Property Get SelTextOnFocus() As Boolean
SelTextOnFocus = m_SelTextFocus
End Property

Public Property Let SelTextOnFocus(ByVal bSel As Boolean)
m_SelTextFocus = bSel
PropertyChanged "SelTextOnFocus"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Lst,Lst,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indica si los elementos de un control se ordenan automáticamente de forma alfabética."
    Sorted = Lst.Sorted
End Property

Public Property Get Text() As String
Text = Txt.Text
End Property

Public Property Let Text(ByVal New_Text As String)
Txt.Text = New_Text
PropertyChanged "Text"
UserControl_Resize
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Txt,Txt,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Devuelve o establece el número de caracteres seleccionados."
  SelLength = Txt.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  Txt.SelLength() = New_SelLength
  PropertyChanged "SelLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Txt,Txt,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Devuelve o establece el punto inicial del texto seleccionado."
  SelStart = Txt.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  Txt.SelStart() = New_SelStart
  PropertyChanged "SelStart"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Txt,Txt,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Devuelve o establece la cadena que contiene el texto seleccionado actualmente."
  SelText = Txt.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
  Txt.SelText() = New_SelText
  PropertyChanged "SelText"
End Property

