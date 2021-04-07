VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl AxBiGrid 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "AxBiGrid.ctx":0000
   ScaleHeight     =   4170
   ScaleWidth      =   8475
   ToolboxBitmap   =   "AxBiGrid.ctx":0041
   Begin AxioGrid.axComboBox cBox 
      Height          =   255
      Left            =   3630
      TabIndex        =   8
      Top             =   3315
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      ItemsInList     =   6
      ItemSelectMode  =   1
      ListIndex       =   -1
      ItemSelectMode  =   1
   End
   Begin VB.CommandButton Boton 
      Caption         =   "<"
      Height          =   255
      Left            =   660
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   230
   End
   Begin VB.TextBox TBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFF4&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1380
      TabIndex        =   6
      Top             =   3255
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtInfoBar 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3585
      Width           =   5865
   End
   Begin VB.ListBox LBox 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   6315
      TabIndex        =   4
      Top             =   3255
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   5040
      TabIndex        =   3
      Top             =   3375
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Divisor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Left            =   3210
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5475
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   45
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   3195
      Index           =   2
      Left            =   3315
      TabIndex        =   2
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   3195
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "AxBiGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''-----------------------------------------------------------
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'''-----------------------------------------------------------

'Extra Events
Public Event BeforeEdit(Cancel As Boolean)
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event CancelEdit(xGrid As eSideGrid, ByVal Row As Long, ByVal Col As Long)
Public Event ValidateEdit(Row As Long, Col As Long, Cancel As Boolean)
Public Event ButtonClick(xGrid As eSideGrid, ByVal Row As Long, ByVal Col As Long)
Public Event CellTextChange(xGrid As eSideGrid, ByVal Row As Long, ByVal Col As Long)
Public Event KeyPressEdit(xGrid As eSideGrid, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
Public Event KeyDownEdit(xGrid As eSideGrid, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

'Mapped Events
Public Event Click(xGrid As eSideGrid, Row As Long, Col As Long)
Public Event Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Public Event DblClick(xGrid As eSideGrid, Row As Long, Col As Long)
Public Event EnterCell()
Public Event LeaveCell()
Public Event RowColChange()
Public Event Scroll()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event SelChange()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event OLEStartDrag(xGrid As eSideGrid, Data As MSFlexGridLib.DataObject, AllowedEffects As Long)
Public Event OLESetData(xGrid As eSideGrid, Data As MSFlexGridLib.DataObject, DataFormat As Integer)
Public Event OLEGiveFeedback(xGrid As eSideGrid, Effect As Long, DefaultCursors As Boolean)
Public Event OLEDragOver(xGrid As eSideGrid, Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEDragDrop(xGrid As eSideGrid, Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(xGrid As eSideGrid, Effect As Long)

Private m_ColsL() As FgCol
Private m_ColsR() As FgCol

Private cFlatScroll(1 To 2) As New cCoolScrollbars
'Private cFlat1               As New cFlatControl
'Private cFlat2               As New cFlatControl

Dim bPrivateCellChange    As Boolean
Dim m_EnterKeyBehaviour   As eEnterkeyBehaviour
Dim m_AutoSizeMode        As eAutoSizeSetting
Dim m_BackColorAlternate  As OLE_COLOR
Dim m_BackColor           As OLE_COLOR
Dim m_BackColorSel        As OLE_COLOR
Dim m_SortMode            As SortSettings
Dim m_AddRows             As Boolean
Dim m_KeyUp               As Integer
Dim m_Editable            As Boolean
Dim m_SplitterFixed       As Boolean
Dim m_SplitterPos         As Long
Dim m_SplitIniPos         As Long
Dim m_SetInfoBar          As eTypeInfoBar
Dim m_txtInfoBar          As String
Dim m_ColEdit             As Boolean
Dim m_Command             As String
Dim ucSH                  As Long
Dim ucSW                  As Long
Dim tPart                 As Long
Dim lRow                  As Long
Dim lCol                  As Long
Dim ObjGrd                As Long
Dim OldRow(1 To 2)        As Integer
Dim lngOldX               As Long
Dim IsMoving              As Boolean
Dim bCboxSel              As Boolean
Dim g                     As Integer

'Dim m_SearchTable As String
'Dim m_SearchField As String

Private Const m_Def_BackColorAlt = &H8000000F

Private sDecimal        As String
Private sThousand       As String
Private sDateDiv        As String
Private sMoney          As String


'************
'New Methods
'************
Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
   Load AxGridAbout
  With AxGridAbout
      .lblProd.Caption = "AxBiGrid "
      .lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
      .lblV1.Caption = App.Major
      .lblV2.Caption = "." & App.Minor
      .lblInfo.Caption = "Dual version of AxGrid: Most Enhanced yet!"
      .Image1.Picture = .Image2.Picture
      .Show (vbModal)
  End With
End Sub

Public Sub AddItem(ByVal eG As eSideGrid, strValue As String, Optional rowIndex As Long)
Dim gG As Integer, i As Long

If eG = eLeftGrid Then
  gG = 2
Else
  gG = 1
End If
  
With fg(eG)
    If Not IsMissing(rowIndex) Then
      .AddItem strValue, rowIndex
      fg(gG).AddItem "", rowIndex
    Else
      .AddItem strValue
      fg(gG).AddItem ""
    End If
End With

For i = 1 To 2
  If m_BackColorAlternate <> m_BackColor Then
     Call SetAlternateRowColors(i, m_BackColorAlternate, m_BackColor)
  Else
     fg(i).BackColor = m_BackColor
  End If
Next i

End Sub

Public Sub AddItemObject(cObj As eTypeControl, strValue As String, Optional oIndex As Long)
Select Case cObj
  Case Is = oListBox
      If Not IsMissing(oIndex) Then
        LBox.AddItem strValue, oIndex
      Else
        LBox.AddItem strValue
      End If
      
  Case Is = oComboBox
      If Not IsMissing(oIndex) Then
        cBox.AddItem strValue, oIndex
      Else
        cBox.AddItem strValue
      End If
      
End Select

End Sub

Public Sub AutoSizeCols(eG As eSideGrid, eFirstCol As Long, eLastCol As Long)
 Dim i As Long, j As Long
 Dim nMaxWidth As Long
 Dim nCurrWidth As Long
                If IsMissing(eLastCol) Then eLastCol = eFirstCol
                Call AutoSizeC(eG, eFirstCol, eLastCol, True)
End Sub

Public Sub ClearGrid(Optional eSide As eSideGrid2)
    If eSide = LeftGrid Then
      fg(1).Clear
    ElseIf eSide = RightGrid Then
      fg(2).Clear
    ElseIf eSide = BothGrids Then
      fg(1).Clear
      fg(2).Clear
    End If

    If m_BackColorAlternate <> m_BackColor Then
       Call SetAlternateRowColors(1, m_BackColorAlternate, m_BackColor)
       Call SetAlternateRowColors(2, m_BackColorAlternate, m_BackColor)
    Else
       fg(1).BackColor = m_BackColor
       fg(2).BackColor = m_BackColor
    End If

End Sub

Public Sub ClearItemObject(cObj As eTypeControl)
Select Case cObj
  Case Is = oListBox
      LBox.Clear
      
  Case Is = oComboBox
      cBox.Clear
      
End Select

End Sub

Public Function CleanValue(sValor As String) As Double

If Left(sValor, 1) = sMoney Then
   Dim sValue As Variant, i As Integer
   For i = 1 To Len(sValor)
      If Not Mid(sValor, i, 1) = sMoney And Not Mid(sValor, i, 1) = sThousand Then
         sValue = sValue & Mid(sValor, i, 1)
      End If
   Next i
Else
  sValue = sValor
End If
CleanValue = Trim(sValue)
Debug.Print sValue

End Function

Private Function fGetLocaleInfo(Valor As RegionalConstant) As String
   Dim Simbolo As String
   Dim r1 As Long
   Dim r2 As Long
   Dim p As Integer
   Dim Locale As Long
     
   Locale = GetUserDefaultLCID()
   r1 = GetLocaleInfo(Locale, Valor, vbNullString, 0)
   'buffer
   Simbolo = String$(r1, 0)
   'En esta llamada devuelve el símbolo en el Buffer
   r2 = GetLocaleInfo(Locale, Valor, Simbolo, r1)
   'Localiza el espacio nulo de la cadena para eliminarla
   p = InStr(Simbolo, Chr$(0))
     
   If p > 0 Then
      'Elimina los nulos
      fGetLocaleInfo = Left$(Simbolo, p - 1)
   End If
     
End Function

Public Sub ExtendLastColumn(eSide As eSideGrid, Optional Col As Long)
    Dim m_lScrollWidth As Long
    m_lScrollWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelY
    
    Dim lCol As Long
    Dim lTotWidth As Long
    Dim lScrollWidth As Long
    Dim nMargin As Long
    nMargin = 95
    
    With fg(eSide)
        ' is there a vertical scrollbar
        lScrollWidth = 0
        If .ScrollBars = flexScrollBarBoth Or .ScrollBars = flexScrollBarBoth Then
            If Not .RowIsVisible(0) Or Not .RowIsVisible(.Rows - 1) Then
                lScrollWidth = m_lScrollWidth
            End If
        End If
    
        Dim nWidth As Long
        nWidth = .Width - lScrollWidth
        Dim nColWidths As Long
        Dim i As Long
        For i = 0 To .Cols - 1
            nColWidths = nColWidths + .ColWidth(i) + .GridLineWidth
        Next
        If .Appearance = flex3D Then
            nMargin = 95
        Else
            nMargin = 35
        End If
        
        Dim nProcessCol As Long
        If Not IsMissing(Col) Then
           nProcessCol = Col
        Else
           nProcessCol = .Cols - 1
        End If
        If nColWidths < nWidth - nMargin Then
            .ColWidth(nProcessCol) = .ColWidth(nProcessCol) + (nWidth - nColWidths - nMargin)
        End If
        If lScrollWidth = 0 Then
        End If
    End With
End Sub

Public Function GridsToHTML() As String
    Dim i As Long
    Dim j As Long
    Dim sText As String
    ' Encabezado
    sText = "<HTML>" & vbCrLf & "<BODY>" & vbCrLf & "<TABLE>" & vbCrLf
    
    ' Leer Grids...
    For i = 0 To fg(1).Rows - 1
      ' Abre Fila HTML
      sText = sText & "<TR>" & vbCrLf
      ' Lee Primer Grid
      For j = 0 To fg(1).Cols - 1
        sText = sText & "<TD>" & fg(1).TextMatrix(i, j) & "</TD>"
      Next
      ' Lee Segundo Grid
      For j = 0 To fg(2).Cols - 1
        sText = sText & "<TD>" & fg(2).TextMatrix(i, j) & "</TD>"
      Next
      
      sText = sText & vbCrLf & "</TR>" & vbCrLf
    Next
    
    ' Cierre HTML
    sText = sText & "</TABLE>" & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"
    
    GridsToHTML = sText

ErrHand:
    Err.Raise Err.Number, Err.Source, Err.Description
    GridsToHTML = ""
End Function

Public Sub RemoveItem(Index As Long)
  fg(1).RemoveItem (Index)
  fg(2).RemoveItem (Index)
End Sub

Public Function SaveAsExcel(sFilename As String)
    Dim myExcel As ExcelFileV2
    Dim lCol2 As Long
    Dim excelDouble As Double
    Dim rowOffset As Long
    Dim aTemp() As String
      
    Set myExcel = New ExcelFileV2
      
    lCol2 = fg(1).Cols

      With myExcel
          .OpenFile sFilename
          ' FlexGrid -> Fixedrows
        For lRow = 1 To fg(1).FixedRows
          ' Primer Grid
          For lCol = 1 To fg(1).Cols
            .EWriteString lRow + rowOffset, lCol, fg(1).TextMatrix(lRow - 1, lCol - 1)
          Next lCol
          ' Segundo Grid
          For lCol = 1 To fg(2).Cols
            .EWriteString lRow + rowOffset, lCol2 + lCol, fg(2).TextMatrix(lRow - 1, lCol - 1)
          Next lCol
        Next lRow
      
      ' Grids Data
        For lRow = fg(1).FixedRows + 1 To fg(1).Rows
          ' FlexGrid -> Fixedcols
          ' Primer Grid
          For lCol = 1 To fg(1).FixedCols
            .EWriteString lRow + rowOffset, lCol, fg(1).TextMatrix(lRow - 1, lCol - 1)
          Next lCol
          ' Segundo Grid
          For lCol = 1 To fg(2).FixedCols
            .EWriteString lRow + rowOffset, lCol2 + lCol, fg(2).TextMatrix(lRow - 1, lCol - 1)
          Next lCol

          ' Primer Grid -> Data
          For lCol = fg(1).FixedCols + 1 To fg(1).Cols
             If IsNumeric(fg(1).TextMatrix(lRow - 1, lCol - 1)) Then
                  excelDouble = CDbl(fg(1).TextMatrix(lRow - 1, lCol - 1)) + 0
                  .EWriteDouble lRow + rowOffset, lCol, excelDouble
             Else
                 .EWriteString lRow + rowOffset, lCol, fg(1).TextMatrix(lRow - 1, lCol - 1)
             End If
          Next lCol
          ' Segundo Grid -> Data
          For lCol = fg(2).FixedCols + 1 To fg(2).Cols
             If IsNumeric(fg(2).TextMatrix(lRow - 1, lCol - 1)) Then
                  excelDouble = CDbl(fg(2).TextMatrix(lRow - 1, lCol - 1)) + 0
                  .EWriteDouble lRow + rowOffset, lCol2 + lCol, excelDouble
             Else
                 .EWriteString lRow + rowOffset, lCol2 + lCol, fg(2).TextMatrix(lRow - 1, lCol - 1)
             End If
          Next lCol
        Next lRow
       
        .CloseFile
      End With
  
End Function

Public Sub SaveAsHTML(FileName As String)
Dim i As Long
    Dim j As Long
    Dim sText As String
    ' Encabezado
    sText = "<HTML>" & vbCrLf & "<BODY>" & vbCrLf & "<TABLE>" & vbCrLf
    
    ' Leer Grids...
    For i = 0 To fg(1).Rows - 1
      ' Abre Fila HTML
      sText = sText & "<TR>" & vbCrLf
      ' Lee Primer Grid
      For j = 0 To fg(1).Cols - 1
        sText = sText & "<TD>" & fg(1).TextMatrix(i, j) & "</TD>"
      Next
      ' Lee Segundo Grid
      For j = 0 To fg(2).Cols - 1
        sText = sText & "<TD>" & fg(2).TextMatrix(i, j) & "</TD>"
      Next
      
      sText = sText & vbCrLf & "</TR>" & vbCrLf
    Next
    
    ' Cierre HTML
    sText = sText & "</TABLE>" & vbCrLf & "</BODY>" & vbCrLf & "</HTML>"
        
    If LenB(FileName) Then
       Dim mFileNum
       mFileNum = FreeFile
       On Error Resume Next
       Kill FileName
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       On Error GoTo ErrHand
       Open FileName For Append As #mFileNum
       Print #mFileNum, sText
       Close #mFileNum
    End If
ErrHand:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub AbortEdit()
    txtEdit.Visible = False
End Sub

Private Function AutoSizeC(eSide As eSideGrid, Optional ByVal lfirstCol As Long = -1, _
                                Optional ByVal lLastCol As Long = -1, Optional bCheckFont As Boolean = False)
  
  Dim lCurCol As Long, lCurRow As Long
  Dim lCellWidth As Long, lColWidth As Long
  Dim bFontBold As Boolean
  Dim dFontSize As Double
  Dim sFontName As String
  Dim myGrid As Object
    
  bPrivateCellChange = True
  If bCheckFont Then
    ' save the forms font settings
    bFontBold = Me.FontBold(eSide)
    sFontName = Me.FontName(eSide)
    dFontSize = Me.FontSize(eSide)
  End If
  
If eSide = eLeftGrid Then
  Set myGrid = fg(1)
Else
  Set myGrid = fg(2)
End If

  With myGrid
    If bCheckFont Then
      lCurRow = .Row
      lCurCol = .Col
    End If
    
    If lfirstCol = -1 Then lfirstCol = 0
    If lLastCol = -1 Then lLastCol = .Cols - 1
    
    For lCol = lfirstCol To lLastCol
      lColWidth = 0
      If bCheckFont Then .Col = lCol
      For lRow = 0 To .Rows - 1
        If bCheckFont Then
          .Row = lRow
          UserControl.FontBold = .CellFontBold
          UserControl.FontName = .CellFontName
          UserControl.FontSize = .CellFontSize
        End If
        lCellWidth = UserControl.TextWidth(.TextMatrix(lRow, lCol))
        If lCellWidth > lColWidth Then lColWidth = lCellWidth
      Next lRow
      .ColWidth(lCol) = lColWidth + UserControl.TextWidth("W")
    Next lCol
    
    If bCheckFont Then
      .Row = lCurRow
      .Col = lCurCol
    End If
  End With
  
  If bCheckFont Then
    ' restore the forms font settings
    UserControl.FontBold = bFontBold
    UserControl.FontName = sFontName
    UserControl.FontSize = dFontSize
  End If
  bPrivateCellChange = False
End Function

Private Sub DisplayFormatedText(eG As eSideGrid, S As String, Row As Long, Col As Long)
'On Error Resume Next
If eG = eLeftGrid Then
    fg(1).TextMatrix(Row, Col) = Format(S, m_ColsL(Col).ColDisplayFormat)
Else
    fg(2).TextMatrix(Row, Col) = Format(S, m_ColsR(Col).ColDisplayFormat)
End If
End Sub

Private Sub EndEdit(Cancel As Boolean)
    Dim nRow As Long
    Dim nCol As Long
    Dim sData
On Error Resume Next
With fg(2)
    sData = Split(txtEdit.Tag, "|")
    nRow = Val(sData(0))
    nCol = Val(sData(1))
    'show temporary
    Dim mOldText As String
    mOldText = .TextMatrix(nRow, nCol)
    bPrivateCellChange = True
    DisplayFormatedText 2, txtEdit.Text, nRow, nCol
    bPrivateCellChange = False
    RaiseEvent ValidateEdit(nRow, nCol, Cancel)
    If Not Cancel Then
        'DisplayFormatedText txtEdit.Text, fg.Row, fg.Col
        DisplayFormatedText 2, txtEdit.Text, nRow, nCol
        RaiseEvent AfterEdit(nRow, nCol)
        txtEdit.Visible = False
        .SetFocus
    Else
       ' If not validated then restore original text
       .TextMatrix(nRow, nCol) = mOldText
    End If
End With
End Sub

Private Sub IsCellVisible()
Dim a As Boolean, b As Boolean
   a = fg(1).CellTop
   b = fg(2).CellTop
End Sub

'*******************
' Internal Required Functions
'****************************
Private Function Max(Val1 As Double, Val2 As Double) As Double
    If Val1 > Val2 Then
       Max = Val1
    Else
       Max = Val2
    End If

End Function

Private Function Min(Val1 As Double, Val2 As Double) As Double
    If Val1 > Val2 Then
       Min = Val2
    Else
       Min = Val1
    End If
End Function

Private Sub MoveCellOnEnter(fGx As Long)
On Error GoTo ErrSub
With fg(fGx)
    If Not m_EnterKeyBehaviour = axEKNone Then
        If m_EnterKeyBehaviour = axEKMoveRight Then
            If .Col < .Cols - 1 Then
               .Col = .Col + 1
            Else
               If .Row < .Rows - 1 Then
                  Dim nRow As Long
                  Dim nCol As Long
                  nCol = .Col
                  nRow = .Row
                  nRow = .Row + 1
                  Dim i As Long
                  For i = .FixedCols To .Cols - 1
                      If .ColWidth(i) > 0 Then
                         nCol = i
                         Exit For
                      End If
                  Next
                  bPrivateCellChange = True
                  .Row = nRow
                  bPrivateCellChange = False
                  If .Col <> nCol Then
                     .Col = nCol
                  End If
               End If
            End If
            
        ElseIf m_EnterKeyBehaviour = axEKMoveDown Then
            If .Row < .Rows - 1 Then .Row = .Row + 1
        End If
        Call IsCellVisible
    End If
End With

ErrSub:
End Sub

Private Sub SetAlternateRowColors(eG As eSideGrid, lColor1 As Long, lColor2 As Long)
    Dim lOrgRow As Long, lOrgCol As Long
    Dim lColor As Long, i As Integer

    bPrivateCellChange = True
      With fg(eG)
          .Redraw = False
          ' save the current cell position
          lOrgRow = .Row
          lOrgCol = .Col
          ' only the data rows
          For lRow = .FixedRows To .Rows - 1
              .Row = lRow
              If lRow / 2 = lRow \ 2 Then
                  lColor = lColor1
              Else
                  lColor = lColor2
              End If
              ' only the data columns
              For lCol = .FixedCols To .Cols - 1
                  .Col = lCol
                  .CellBackColor = lColor
              Next lCol
          Next lRow
          ' restore the orginal cell position
          .Row = lOrgRow
          .Col = lOrgCol
          .Redraw = True
      End With
    
    bPrivateCellChange = False
End Sub

Private Sub SetColumnObject(eG As Integer, lCol As Long)
cBox.Visible = False
LBox.Visible = False
TBox.Visible = False
Boton.Visible = False

ObjGrd = eG
fg(eG).Col = lCol

If eG = 1 Then
  Select Case m_ColsL(lCol).ColType
    Case Is = eComboBoxColumn
      GoTo SetCBoxH
    Case Is = eListBoxColumn
      GoTo SetLBoxH
    Case Is = eTextBoxColumn
      GoTo SetTBoxH
    Case Is = eButtonColumn
      GoTo SetButtonH
  End Select
ElseIf eG = 2 Then
  Select Case m_ColsR(lCol).ColType
    Case Is = eComboBoxColumn
      GoTo SetCBoxH
    Case Is = eListBoxColumn
      GoTo SetLBoxH
    Case Is = eTextBoxColumn
      GoTo SetTBoxH
    Case Is = eButtonColumn
      GoTo SetButtonH
  End Select
End If

Exit Sub

On Error Resume Next

SetCBoxH:
    With fg(eG)
        If .Col < lCol Then Exit Sub
        cBox.Move .Left + .CellLeft - 8, .Top + .CellTop - 8, .CellWidth
    End With
    'If cBox.ListCount > 1 Then cBox.ListIndex = 0
    cBox.Text = fg(eG).Text
    cBox.Visible = True
    cBox.SetFocus
    Exit Sub
    
SetLBoxH:
    With fg(eG)
        If .Col < lCol Then Exit Sub
        LBox.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, 1350
    End With
    'If LBox.ListCount > 1 Then LBox.ListIndex = 0
    LBox.Text = fg(eG).Text
    LBox.Visible = True
    LBox.SetFocus
    Exit Sub
    
SetTBoxH:
    With fg(eG)
        If .Col < lCol Then Exit Sub
        TBox.Move .Left + .CellLeft + 3, .Top + .CellTop + 3, .CellWidth - 11, .CellHeight - 11
        TBox.Text = .Text
    End With
    TBox.Visible = True
    TBox.SelStart = 0
    TBox.SelLength = Len(TBox)
    TBox.SetFocus
    Exit Sub

SetButtonH:
    With fg(eG)
        If .Col < lCol Then Exit Sub
        Boton.Move .Left + .CellLeft + .CellWidth, .Top + .CellTop, 230, .CellHeight
    End With
    Boton.Visible = True
    Exit Sub
End Sub

Private Sub StartKeyEdit(KeyAscii As Integer, Optional bShowOldText As Boolean)
    If Not m_Editable Then Exit Sub
  
  With fg(2)
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft - 2, .Top + .CellTop - 2, .CellWidth - 8, .CellHeight - 8
    
    Dim Cancel As Boolean
    RaiseEvent BeforeEdit(Cancel)
    
    If Not Cancel Then
        If bShowOldText Then
           txtEdit.Text = .Text
           txtEdit.SelStart = 0
           txtEdit.SelLength = Len(txtEdit.Text)
        Else
           txtEdit.Text = Chr$(KeyAscii)
           txtEdit.SelStart = 1
        End If
        txtEdit.Tag = .Row & "|" & .Col
        txtEdit.Visible = True
        txtEdit.SetFocus
    End If
  End With
  
End Sub

Private Sub Boton_Click()
With fg(ObjGrd)
  RaiseEvent ButtonClick(ObjGrd, .Row, .Col)
End With
End Sub

Private Sub cBox_Click()
'On Error Resume Next
'fg(ObjGrd).Text = cBox.Text
'cBox.Visible = False
bCboxSel = True
End Sub

Private Sub cBox_DblClick()
On Error Resume Next
With fg(ObjGrd)
  .Text = cBox.Text
  cBox.Visible = False
  RaiseEvent CellTextChange(ObjGrd, .Row, .Col)
End With
End Sub

Private Sub cBox_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
  If bCboxSel = True Then
    With fg(ObjGrd)
      .Text = cBox.Text
      cBox.Visible = False
      RaiseEvent CellTextChange(ObjGrd, .Row, .Col)
    End With
  Else
    cBox.Visible = False
  End If
End If
End Sub

Private Sub cBox_KeyUp(KeyCode As Integer, Shift As Integer)
bCboxSel = AutoComplete(cBox, KeyCode)
If KeyCode = vbKeyEscape Then
  cBox.Visible = False
  RaiseEvent CancelEdit(ObjGrd, fg(ObjGrd).Row, fg(ObjGrd).Col)
End If

End Sub

Private Sub Divisor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  lngOldX = x
  IsMoving = True
  Divisor.Height = UserControl.ScaleHeight
  Divisor.Top = 0 'UserControl.ScaleHeight * -1
  
  fg(1).Redraw = False
  fg(2).Redraw = False
  cBox.Visible = False
  TBox.Visible = False
  LBox.Visible = False
  Boton.Visible = False

End Sub

Private Sub Divisor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If m_SplitterFixed = True Then Exit Sub
If Button = vbLeftButton And IsMoving = True Then 'And m_Ajustable = True Then
'  Dim res As Long
'  Call ReleaseCapture
'  res = SendMessage(Divisor.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Divisor.Left = Divisor.Left - (lngOldX - x)
    
  ' Ajusto Ancho xGrid1
  fg(1).Width = Divisor.Left
  ' Ajusto Ancho xGrid2
  fg(2).Left = fg(1).Width + Divisor.Width
  fg(2).Width = ucSW - fg(2).Left
  ' Ajusto Top Divisor
  Divisor.Top = 0
  Divisor.Height = UserControl.ScaleHeight - 270
End If

End Sub

Private Sub Divisor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  IsMoving = False
  fg(1).Redraw = True
  fg(2).Redraw = True

  m_SplitterPos = Divisor.Left
  PropertyChanged "SplitterPos"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' * * * '''''''''''''''''''''''''''''''''''''''''''
Private Sub fg_Click(Index As Integer)
On Error GoTo ErrSub0

Select Case Index
  Case Is = 1
    ' Desmarco Anterior
    If m_BackColorAlternate <> m_BackColor Then
       Call SetAlternateRowColors(2, m_BackColorAlternate, m_BackColor)
    Else
       fg(2).BackColor = m_BackColor
    End If
    
    fg(2).Row = fg(1).Row
    
    For g = 0 To fg(2).Cols - 1
      fg(2).Col = g
      fg(2).CellBackColor = m_BackColorSel
    Next g
  
    lRow = fg(1).Row
    lCol = fg(1).Col
    
  Case Is = 2
    ' Desmarco Anterior
    If m_BackColorAlternate <> m_BackColor Then
       Call SetAlternateRowColors(1, m_BackColorAlternate, m_BackColor)
    Else
       fg(1).BackColor = m_BackColor
    End If
    
    fg(1).Row = fg(2).Row
    
    For g = 1 To fg(1).Cols - 1
      fg(1).Col = g
      fg(1).CellBackColor = m_BackColorSel
    Next g
  
    lRow = fg(2).Row
    lCol = fg(2).Col
    
End Select

With txtInfoBar
  Dim sRowText As String
  .Text = "Left: [" & lRow & ":" & fg(1).Col & "] Right: [" & lRow & ":" & fg(2).Col & "] "
  sRowText = ""
  
  Select Case m_SetInfoBar
    Case Is = RightGridInfo
      ' Recorrer Todo el Grid
      For g = 0 To fg(2).Cols - 1
        sRowText = sRowText & fg(2).TextMatrix(lRow, g) & "|"
      Next g
      .Text = .Text & Mid$(sRowText, 1, Len(sRowText) - 1)
      
    Case Is = LeftGridInfo
      ' Recorrer Todo el Grid
      For g = 0 To fg(1).Cols - 1
        sRowText = sRowText & fg(1).TextMatrix(lRow, g) & "|"
      Next g
      .Text = .Text & Mid$(sRowText, 1, Len(sRowText) - 1)
    
    Case Is = BothGridsInfo
      ' Recorrer Todo el Grid
      For g = 0 To fg(1).Cols - 1
        sRowText = sRowText & fg(1).TextMatrix(lRow, g) & "|"
      Next g
      ' First Grid
      .Text = .Text & Mid$(sRowText, 1, Len(sRowText) - 1)
      sRowText = ""
      ' Recorrer Todo el Grid
      For g = 0 To fg(2).Cols - 1
        sRowText = sRowText & fg(2).TextMatrix(lRow, g) & "|"
      Next g
      ' Second Grid
      .Text = .Text & Mid$(sRowText, 1, Len(sRowText) - 1)
      
    Case Is = CustomInfo
      '.Text = m_txtInfoBar
      'PropertyChanged "InfoBarText"
  End Select
End With


RaiseEvent Click(CLng(Index), lRow, lCol)
Exit Sub
ErrSub0:
RaiseEvent Click(CLng(Index), lRow, lCol)
MsgBox Err.Number & " " & Err.Description, vbOKOnly
End Sub

Private Sub fg_Compare(Index As Integer, ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    RaiseEvent Compare(Row1, Row2, Cmp)

End Sub

Private Sub fg_DblClick(Index As Integer)
Static bTipo As Boolean

With fg(Index)
    If (.MouseRow = 0) Then
      ' Ordena en forma ascendente
      If bTipo Then
        .Col = .MouseCol
        .Sort = 2
        bTipo = False
        ' Ordena en forma descendente
      Else
        .Col = .MouseCol
        .Sort = 1
        bTipo = True
      End If
    End If
End With

    RaiseEvent DblClick(CLng(Index), fg(Index).Row, fg(Index).Col)
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If Not bPrivateCellChange Then
       RaiseEvent EnterCell
    End If
End Sub

Private Sub fg_GotFocus(Index As Integer)

If m_BackColorAlternate <> m_BackColor Then
   Call SetAlternateRowColors(CLng(Index), m_BackColorAlternate, m_BackColor)
Else
   fg(Index).BackColor = m_BackColor
End If

'If Index = 2 Then
    If txtEdit.Visible = True Then
        Dim Cancel As Boolean
        Call EndEdit(Cancel)
        If Cancel Then
           txtEdit.SetFocus
        End If
    End If
'End If

HookScroll fg(Index)

End Sub

Private Sub fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
If m_Editable = True Then
  If KeyCode = vbKeyF2 Then
      Call StartKeyEdit(0, True)
  ElseIf KeyCode = vbKeyDelete Then
      fg(Index).Text = ""
'    If fg(Index).Row >= fg(Index).FixedRows Then
'      If MsgBox("La Linea " & fg(Index).Row & " se eliminará!" & vbNewLine & _
'                        "_____________Esta seguro?", vbExclamation + vbYesNo, "Eliminar Linea") = vbYes Then
'        'Delete the row
'        fg(1).RemoveItem (fg(1).Row)
'        fg(2).RemoveItem (fg(1).Row)
'        'prevent beep
'        KeyCode = 0
'      End If
'    End If
  End If
End If
End Sub

Private Sub fg_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim sInputMask As String
    Select Case Index
        Case 1
            sInputMask = m_ColsL(fg(1).Col).ColInputMask
        Case 2
            sInputMask = m_ColsR(fg(2).Col).ColInputMask
    End Select
    
    If LenB(sInputMask) > 0 Then
       If Not KeyAscii = vbKeyReturn Then
           txtEdit.Text = ""
           NumKeyPress KeyAscii, txtEdit, sInputMask
           If KeyAscii = 0 Then
              Exit Sub
           End If
       End If
    End If
    
    RaiseEvent KeyPress(KeyAscii)
    
    If KeyAscii = 0 Then Exit Sub
    
    RaiseEvent KeyPressEdit(CLng(Index), fg(Index).Row, fg(Index).Col, KeyAscii)
    
    If KeyAscii = vbKeyTab Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        Call MoveCellOnEnter(CLng(Index))
        KeyAscii = 0
    End If
    
    If KeyAscii > 0 Then
        Call StartKeyEdit(KeyAscii)
    End If
End Sub

Private Sub fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case m_EnterKeyBehaviour
   Case Is = axEKMoveDown
        If fg(Index).Row = fg(Index).Rows - 1 Then
          If m_AddRows = True And KeyCode = vbKeyReturn And m_KeyUp = 2 Then
            fg(1).Rows = fg(1).Rows + 1
            fg(2).Rows = fg(1).Rows
            m_KeyUp = 0
          End If
          
          m_KeyUp = m_KeyUp + 1
        End If
        
   Case Is = axEKMoveRight
        If fg(Index).Row = fg(Index).Rows - 1 And fg(Index).Col = fg(Index).Cols - 1 Then
          If m_AddRows = True And KeyCode = vbKeyReturn And m_KeyUp = 2 Then
            fg(1).Rows = fg(1).Rows + 1
            fg(2).Rows = fg(1).Rows
            m_KeyUp = 0
          End If
          
          m_KeyUp = m_KeyUp + 1
        End If

End Select
    
   Call SetAlternateRowColors(1, m_BackColorAlternate, m_BackColor)
   Call SetAlternateRowColors(2, m_BackColorAlternate, m_BackColor)
 
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub fg_LeaveCell(Index As Integer)
    If Not bPrivateCellChange Then
       RaiseEvent LeaveCell
    End If
End Sub

Private Sub fg_LostFocus(Index As Integer)
OldRow(Index) = fg(Index).Row

'cBox.Visible = False
'TBox.Visible = False
'LBox.Visible = False
'Boton.Visible = False

UnHookScroll fg(Index)
End Sub

Private Sub fG_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

  With fg(Index)
    Call SetColumnObject(Index, .Col)
  End With

   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub fg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub fg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Static CurrentWidth As Single

  ' Check to see if the Cell's width has changed.
  If fg(Index).CellWidth <> CurrentWidth Then
    cBox.Width = fg(Index).CellWidth
    LBox.Width = fg(Index).CellWidth + 270
    CurrentWidth = fg(Index).CellWidth
  End If

   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub fg_OLECompleteDrag(Index As Integer, Effect As Long)
   RaiseEvent OLECompleteDrag(CLng(Index), Effect)
End Sub

Private Sub fg_OLEDragDrop(Index As Integer, Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent OLEDragDrop(CLng(Index), Data, Effect, Button, Shift, x, y)
End Sub

Private Sub fg_OLEDragOver(Index As Integer, Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   RaiseEvent OLEDragOver(CLng(Index), Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub fg_OLEGiveFeedback(Index As Integer, Effect As Long, DefaultCursors As Boolean)
   RaiseEvent OLEGiveFeedback(CLng(Index), Effect, DefaultCursors)
End Sub

Private Sub fg_OLESetData(Index As Integer, Data As MSFlexGridLib.DataObject, DataFormat As Integer)
   RaiseEvent OLESetData(CLng(Index), Data, DataFormat)
End Sub

Private Sub fg_OLEStartDrag(Index As Integer, Data As MSFlexGridLib.DataObject, AllowedEffects As Long)
   RaiseEvent OLEStartDrag(CLng(Index), Data, AllowedEffects)
End Sub

Private Sub fg_RowColChange(Index As Integer)
   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub fg_Scroll(Index As Integer)
  If txtEdit.Visible = True Then
    AbortEdit
  End If
  
  cBox.Visible = False
  LBox.Visible = False
  TBox.Visible = False
  
  On Error Resume Next
  If Index = 1 Then
    fg(2).TopRow = fg(1).TopRow
  Else
    fg(1).TopRow = fg(2).TopRow
  End If
  RaiseEvent Scroll
End Sub

Private Sub fg_SelChange(Index As Integer)

   If Not bPrivateCellChange Then
      RaiseEvent SelChange
   End If
End Sub

Private Sub LBox_DblClick()
On Error Resume Next
fg(ObjGrd).Text = LBox.Text
LBox.Visible = False
RaiseEvent CellTextChange(ObjGrd, fg(ObjGrd).Row, fg(ObjGrd).Col)
End Sub

Private Sub LBox_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
  fg(ObjGrd).Text = LBox.Text
  LBox.Visible = False
  RaiseEvent CellTextChange(ObjGrd, fg(ObjGrd).Row, fg(ObjGrd).Col)
ElseIf KeyAscii = vbKeyEscape Then
  LBox.Visible = False
  RaiseEvent CancelEdit(ObjGrd, fg(ObjGrd).Row, fg(ObjGrd).Col)
End If
End Sub

Private Sub TBox_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
  With fg(ObjGrd)
    .Text = TBox.Text
    TBox.Visible = False
    KeyAscii = 0
    RaiseEvent CellTextChange(ObjGrd, .Row, .Col)
  End With
ElseIf KeyAscii = vbKeyEscape Then
  TBox.Visible = False
    RaiseEvent CancelEdit(ObjGrd, fg(ObjGrd).Row, fg(ObjGrd).Col)
End If

End Sub

''''''Right Grid Edit''''''''''''''''''''
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sInputMask As String
With fg(2)
    sInputMask = m_ColsR(.Col).ColInputMask
    If LenB(sInputMask) > 0 Then
       NumKeyDown KeyCode, txtEdit, sInputMask
    End If
    
    RaiseEvent KeyDownEdit(2, .Row, .Col, KeyCode, Shift)
    Dim Cancel As Boolean
    
    If KeyCode = vbKeyDown Then
        Call EndEdit(Cancel)
        If Not Cancel Then
           If .Row < .Rows - 1 Then .Row = .Row + 1
        End If
    End If
    If KeyCode = vbKeyUp Then
        Call EndEdit(Cancel)
        If Not Cancel Then
            If .Row > .FixedRows Then .Row = .Row - 1
        End If
    End If
End With
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim sInputMask As String
    
With fg(2)
    sInputMask = m_ColsR(.Col).ColInputMask
    If LenB(sInputMask) > 0 Then
       If Not KeyAscii = vbKeyReturn Then
           NumKeyPress KeyAscii, txtEdit, sInputMask
           If KeyAscii = 0 Then
              Exit Sub
           End If
       End If
    End If
    
    RaiseEvent KeyPressEdit(2, fg(2).Row, fg(2).Col, KeyAscii) '(.Row, .Col, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        Dim Cancel As Boolean
        Call EndEdit(Cancel)
        If Not Cancel Then Call MoveCellOnEnter(2)
    End If
End With

    If KeyAscii = vbKeyEscape Then
        AbortEdit
    End If
End Sub

Private Sub txtEdit_LostFocus()
  Dim Cancel As Boolean
  Call EndEdit(Cancel)
End Sub
''''''''''''''''''''''''''''''''''

Private Sub UserControl_Initialize()
    ReDim m_ColsR(fg(2).Cols)
    ReDim m_ColsL(fg(1).Cols)

    'cFlat1.Attach cBox
    'cFlat2.Attach LBox
    
'   AttachMessage Me, fg.hwnd, WM_MOUSEMOVE

End Sub

Private Sub UserControl_InitProperties()
m_BackColorAlternate = m_Def_BackColorAlt
m_SplitIniPos = UserControl.ScaleWidth / 3
m_SetInfoBar = BothGrids
End Sub

'*********************************
'Load property values from storage
'*********************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Integer
    ' New Properties
    m_AddRows = PropBag.ReadProperty("AddRowsOnDemand", False)
    m_SetInfoBar = PropBag.ReadProperty("SetInfoBar", BothGrids)
    m_SplitterPos = PropBag.ReadProperty("SplitterPos", m_SplitIniPos)
    m_SplitterFixed = PropBag.ReadProperty("SplitterFixed", False)
    m_EnterKeyBehaviour = PropBag.ReadProperty("EnterKeyBehaviour", axEKMoveRight)
    m_Editable = PropBag.ReadProperty("Editable", False)
    m_AutoSizeMode = PropBag.ReadProperty("AutoSizeMode", axAutoSizeColWidth)
    m_SortMode = PropBag.ReadProperty("SortColumnMode", flexSortNone)
    m_BackColorAlternate = PropBag.ReadProperty("BackColorAlternate", m_Def_BackColorAlt)
    txtInfoBar.Visible = PropBag.ReadProperty("ShowInfoBar", True)
    ' Mapped Properties
    ColsLeft = PropBag.ReadProperty("ColsLeft", 2)
    ColsRight = PropBag.ReadProperty("ColsRight", 2)

For i = 1 To 2
    fg(i).GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", flexGridInset)
    fg(i).GridLines = PropBag.ReadProperty("GridLines", flexGridFlat)
    fg(i).AllowBigSelection = PropBag.ReadProperty("AllowBigSelection", True)
    fg(i).AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 0)
    fg(i).Appearance = PropBag.ReadProperty("Appearance", 1)
    fg(i).BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    'Set Edit's backcolor similar to grid's backcolor
    txtEdit.BackColor = fg(i).BackColor
    fg(i).BackColorBkg = PropBag.ReadProperty("BackColorBkg", &H808080)
    fg(i).BackColorFixed = PropBag.ReadProperty("BackColorFixed", &H8000000F)
    
    fg(i).BackColorSel = PropBag.ReadProperty("BackColorSel", &H8000000D)
    m_BackColorSel = PropBag.ReadProperty("BackColorSel", &H8000000D)
    
    fg(i).BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    'fg(i).Cols = PropBag.ReadProperty("Cols", 2)
    fg(i).Enabled = PropBag.ReadProperty("Enabled", True)
    fg(i).FillStyle = PropBag.ReadProperty("FillStyle", 0)
    fg(i).FixedCols = PropBag.ReadProperty("FixedCols", 1)
    fg(i).FixedRows = PropBag.ReadProperty("FixedRows", 1)
    fg(i).FocusRect = PropBag.ReadProperty("FocusRect", 1)
    Set fg(i).Font = PropBag.ReadProperty("Font", Ambient.Font)
    fg(i).ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    fg(i).ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", &H80000012)
    fg(i).ForeColorSel = PropBag.ReadProperty("ForeColorSel", &H8000000E)
    fg(i).FormatString = PropBag.ReadProperty("FormatString", "")
    fg(i).GridColor = PropBag.ReadProperty("GridColor", &HC0C0C0)
    fg(i).GridColorFixed = PropBag.ReadProperty("GridColorFixed", &H0&)
    fg(i).GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
    fg(i).HighLight = PropBag.ReadProperty("HighLight", 1)
    fg(i).MergeCells = PropBag.ReadProperty("MergeCells", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    fg(i).MousePointer = PropBag.ReadProperty("MousePointer", 0)
    fg(i).OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    fg(i).PictureType = PropBag.ReadProperty("PictureType", 0)
    fg(i).Redraw = PropBag.ReadProperty("Redraw", True)
    fg(i).RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    fg(i).RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
    fg(i).Rows = PropBag.ReadProperty("Rows", 5)
    fg(i).ScrollBars = PropBag.ReadProperty("ScrollBars", 3)
    fg(i).ScrollTrack = PropBag.ReadProperty("ScrollTrack", True)
    fg(i).SelectionMode = PropBag.ReadProperty("SelectionMode", 0)
    fg(i).Sort = PropBag.ReadProperty("Sort", 0)
    fg(i).TextStyle = PropBag.ReadProperty("TextStyle", 0)
    fg(i).TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", 0)
    fg(i).WordWrap = PropBag.ReadProperty("WordWrap", False)
    
    Set txtEdit.Font = fg(i).Font
    
    If m_BackColorAlternate <> fg(i).BackColor Then
       Call SetAlternateRowColors(CLng(i), m_BackColorAlternate, m_BackColor)
    End If
Next i

    sDecimal = fGetLocaleInfo(LOCALE_SDECIMAL)
    sThousand = fGetLocaleInfo(LOCALE_SMONTHOUSANDSEP)
    sDateDiv = fGetLocaleInfo(LOCALE_SDATE)
    sMoney = fGetLocaleInfo(LOCALE_SCURRENCY)

End Sub

Private Sub UserControl_Resize()
ucSH = UserControl.ScaleHeight
ucSW = UserControl.ScaleWidth

If m_SplitterPos > ucSW Then
  m_SplitterPos = ucSW - 1000
ElseIf m_SplitterPos = 0 Then
  m_SplitterPos = ucSW / 3
End If

Divisor.Left = m_SplitterPos
Divisor.Height = ucSH - 270
Divisor.Top = 0

With txtInfoBar
  .Top = ucSH - .Height + 1
  .Left = 0
  .Width = ucSW
End With

With fg(1)
  .Top = 0
  .Left = 0
  .Width = m_SplitterPos
  ' Setting Rows MinHeight
  .RowHeightMin = cBox.Height
End With

With fg(2)
  .Top = 0
  .Left = m_SplitterPos + Divisor.Width
  .Width = ucSW - .Left
  ' Setting Rows MinHeight
  .RowHeightMin = cBox.Height
End With

If txtInfoBar.Visible = True Then
      fg(1).Height = ucSH - 270
      fg(2).Height = ucSH - 270
Else
      fg(1).Height = ucSH
      fg(2).Height = ucSH
End If

End Sub

Private Sub UserControl_Show()
fg(2).FixedCols = 0
With fg(1)
  .FixedCols = 1
  .ColWidth(0) = 300
End With
UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
''   DetachMessage Me, fg.hwnd, WM_MOUSEMOVE
'  If IsHooked Then
'    Unhook   ' Stop checking messages.
'  End If
UnHookScroll fg(1)
UnHookScroll fg(2)
End Sub

'********************************
'Write property values to storage
'********************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Integer
    ' New Properties
    Call PropBag.WriteProperty("AddRowsOnDemand", m_AddRows, False)
    Call PropBag.WriteProperty("SetInfoBar", m_SetInfoBar, BothGrids)
    Call PropBag.WriteProperty("SplitterPos", m_SplitterPos, m_SplitIniPos)
    Call PropBag.WriteProperty("SplitterFixed", m_SplitterFixed, False)
    Call PropBag.WriteProperty("EnterKeyBehaviour", m_EnterKeyBehaviour, axEKMoveRight)
    Call PropBag.WriteProperty("Editable", m_Editable, False)
    Call PropBag.WriteProperty("AutoSizeMode", m_AutoSizeMode, axAutoSizeColWidth)
    Call PropBag.WriteProperty("SortColumnMode", m_SortMode, flexSortNone)
    Call PropBag.WriteProperty("BackColorAlternate", m_BackColorAlternate, m_Def_BackColorAlt)
    Call PropBag.WriteProperty("ShowInfoBar", txtInfoBar.Visible, True)
    ' Mapped Properties
    Call PropBag.WriteProperty("ColsLeft", fg(1).Cols, 2)
    Call PropBag.WriteProperty("ColsRight", fg(2).Cols, 2)
    
For i = 1 To 2
    Call PropBag.WriteProperty("GridLines", fg(i).GridLines, 1)
    Call PropBag.WriteProperty("GridLinesFixed", fg(i).GridLinesFixed, 1)
    Call PropBag.WriteProperty("AllowBigSelection", fg(i).AllowBigSelection, True)
    Call PropBag.WriteProperty("AllowUserResizing", fg(i).AllowUserResizing, 0)
    Call PropBag.WriteProperty("Appearance", fg(i).Appearance, 1)
    Call PropBag.WriteProperty("BackColor", fg(i).BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColorBkg", fg(i).BackColorBkg, &H808080)
    Call PropBag.WriteProperty("BackColorFixed", fg(i).BackColorFixed, &H8000000F)
    Call PropBag.WriteProperty("BackColorSel", fg(i).BackColorSel, &H8000000D)
    Call PropBag.WriteProperty("BorderStyle", fg(i).BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", fg(i).Enabled, True)
    Call PropBag.WriteProperty("FillStyle", fg(i).FillStyle, 0)
    Call PropBag.WriteProperty("FixedCols", fg(i).FixedCols, 1)
    Call PropBag.WriteProperty("FixedRows", fg(i).FixedRows, 1)
    Call PropBag.WriteProperty("FocusRect", fg(i).FocusRect, 1)
    Call PropBag.WriteProperty("Font", fg(i).Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", fg(i).ForeColor, &H80000008)
    Call PropBag.WriteProperty("ForeColorFixed", fg(i).ForeColorFixed, &H80000012)
    Call PropBag.WriteProperty("ForeColorSel", fg(i).ForeColorSel, &H8000000E)
    Call PropBag.WriteProperty("FormatString", fg(i).FormatString, "")
    Call PropBag.WriteProperty("GridColor", fg(i).GridColor, &HC0C0C0)
    Call PropBag.WriteProperty("GridColorFixed", fg(i).GridColorFixed, &H0&)
    Call PropBag.WriteProperty("GridLineWidth", fg(i).GridLineWidth, 1)
    Call PropBag.WriteProperty("HighLight", fg(i).HighLight, 1)
    Call PropBag.WriteProperty("MergeCells", fg(i).MergeCells, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", fg(i).MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", fg(i).OLEDropMode, 0)
    Call PropBag.WriteProperty("PictureType", fg(i).PictureType, 0)
    Call PropBag.WriteProperty("Redraw", fg(i).Redraw, True)
    Call PropBag.WriteProperty("RightToLeft", fg(i).RightToLeft, False)
    Call PropBag.WriteProperty("RowHeightMin", fg(i).RowHeightMin, 0)
    Call PropBag.WriteProperty("Rows", fg(i).Rows, 5)
    Call PropBag.WriteProperty("ScrollBars", fg(i).ScrollBars, 3)
    Call PropBag.WriteProperty("ScrollTrack", fg(i).ScrollTrack, True)
    Call PropBag.WriteProperty("SelectionMode", fg(i).SelectionMode, 0)
    Call PropBag.WriteProperty("TextStyle", fg(i).TextStyle, 0)
    Call PropBag.WriteProperty("TextStyleFixed", fg(i).TextStyleFixed, 0)
    Call PropBag.WriteProperty("WordWrap", fg(i).WordWrap, False)
Next i
End Sub

'**************************************
' Properties Mapped to Flexgrid Control
'**************************************

Public Property Get AddRowsOnDemand() As Boolean
  AddRowsOnDemand = m_AddRows
End Property

Public Property Let AddRowsOnDemand(new_AddRows As Boolean)
  m_AddRows = new_AddRows
  PropertyChanged "AddRowsOnDemand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,AllowBigSelection
'Public Property Get AllowBigSelection(eSide As eSideGrid) As Boolean
'    AllowBigSelection = fg(2).AllowBigSelection
'End Property

Public Property Let AllowBigSelection(eSide As eSideGrid, ByVal New_AllowBigSelection As Boolean)
Select Case eSide
  Case Is = eLeftGrid
    fg(1).AllowBigSelection() = New_AllowBigSelection
  Case Is = eRightGrid
    fg(2).AllowBigSelection() = New_AllowBigSelection
End Select

    PropertyChanged "AllowBigSelection"
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=fg,fg,-1,AllowUserResizing
'Public Property Get AllowUserResizing(eSide As eSideGrid) As AllowUserResizeSettings
'    AllowUserResizing = fg(2).AllowUserResizing
'End Property

Public Property Let AllowUserResizing(eSide As eSideGrid, ByVal New_AllowUserResizing As AllowUserResizeSettings)
Select Case eSide
  Case Is = eLeftGrid
    fg(1).AllowUserResizing() = New_AllowUserResizing
  Case Is = eRightGrid
    fg(2).AllowUserResizing() = New_AllowUserResizing
End Select

   PropertyChanged "AllowUserResizing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Appearance
Public Property Get Appearance() As AppearanceSettings
    Appearance = fg(1).Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceSettings)
If New_Appearance = flex3D Then
    fg(1).Appearance() = 1
    fg(2).Appearance() = 1
    txtInfoBar.Appearance = 1
    cFlatScroll(1).UninitializeCoolSB
    cFlatScroll(2).UninitializeCoolSB
Else
    fg(1).Appearance() = 0
    fg(2).Appearance() = 0
    txtInfoBar.Appearance = 0
    cFlatScroll(1).InitializeCoolSB fg(1).hwnd, False
    cFlatScroll(2).InitializeCoolSB fg(2).hwnd, False
End If

    PropertyChanged "Appearance"
End Property

Public Property Get AutoSizeMode() As eAutoSizeSetting
    AutoSizeMode = m_AutoSizeMode
End Property

Public Property Let AutoSizeMode(ByVal New_AutoSizeMode As eAutoSizeSetting)
    m_AutoSizeMode = New_AutoSizeMode
    PropertyChanged "AutoSizeMode"
End Property

Public Property Get BackColorAlternate() As OLE_COLOR
    BackColorAlternate = m_BackColorAlternate
End Property

Public Property Let BackColorAlternate(ByVal New_BackColorAlternate As OLE_COLOR)
    m_BackColorAlternate = New_BackColorAlternate
    If m_BackColorAlternate <> fg(1).BackColor Then
       Call SetAlternateRowColors(1, m_BackColorAlternate, fg(1).BackColor)
       Call SetAlternateRowColors(2, m_BackColorAlternate, fg(2).BackColor)
    End If
    PropertyChanged "BackColorAlternate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColorBkg
Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = fg(1).BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    fg(1).BackColorBkg() = New_BackColorBkg
    fg(2).BackColorBkg() = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColorFixed
Public Property Get BackColorFixed() As OLE_COLOR
    BackColorFixed = fg(1).BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
    fg(1).BackColorFixed() = New_BackColorFixed
    fg(2).BackColorFixed() = New_BackColorFixed
    PropertyChanged "BackColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = fg(1).BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If txtEdit.BackColor = fg(1).BackColor Then
       txtEdit.BackColor = New_BackColor
    End If
    
    m_BackColor = New_BackColor
    fg(1).BackColor = New_BackColor
    fg(2).BackColor = New_BackColor
    txtInfoBar.BackColor = New_BackColor
    
    PropertyChanged "BackColor"
    
    Dim OldBackColorAlternate As Long
    OldBackColorAlternate = m_BackColorAlternate
    BackColorAlternate = New_BackColor
    If m_BackColorAlternate <> OldBackColorAlternate Then
       Call SetAlternateRowColors(1, m_BackColorAlternate, m_BackColorAlternate)
       Call SetAlternateRowColors(2, m_BackColorAlternate, m_BackColorAlternate)
    End If

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColorSel
Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = fg(1).BackColorSel
End Property

Public Property Let BackColorSel(ByVal New_BackColorSel As OLE_COLOR)
    m_BackColorSel = New_BackColorSel
    fg(1).BackColorSel() = New_BackColorSel
    fg(2).BackColorSel() = New_BackColorSel
    PropertyChanged "BackColorSel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = fg(1).BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    fg(1).BorderStyle() = New_BorderStyle
    fg(2).BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get CalculateColumn(eSide As eSideGrid, Settings As eSubTotalSettings, eColumn As Long, _
                                    eRowInitial As Long, eRowFinal As Long) As Double
         bPrivateCellChange = True
         Dim nValue As Double
         nValue = 0
         Dim bFirst As Boolean
         bFirst = True
         Dim i As Long
         Dim iDX As Integer
         
    If eSide = eLeftGrid Then
      iDX = 1
    Else
      iDX = 2
    End If
        
    With fg(iDX)
         For i = eRowInitial To eRowFinal
            Select Case Settings
                Case Is = axSTSum
                     nValue = nValue + CleanValue(.TextMatrix(i, eColumn))
                Case Is = axSTMax
                     If bFirst Then
                        nValue = CleanValue(.TextMatrix(i, eColumn))
                        bFirst = False
                     Else
                        nValue = Max(nValue, CleanValue(.TextMatrix(i, eColumn)))
                     End If
                Case Is = axSTMin
                     If bFirst Then
                        nValue = CleanValue(.TextMatrix(i, eColumn))
                        bFirst = False
                     Else
                        nValue = Min(nValue, CleanValue(.TextMatrix(i, eColumn)))
                     End If
                Case Is = axSTCount
                        nValue = nValue + IIf(Len(.TextMatrix(i, eColumn)) > 0, 1, 0)
            End Select
         Next i
         
         If Settings = axSTMultiply Then
            nValue = Val(.TextMatrix(eRowInitial, eColumn)) * CleanValue(.TextMatrix(eRowFinal, eColumn))
         End If
    End With
    
         bPrivateCellChange = False
        CalculateColumn = nValue
End Property

Public Property Get CalculateMatrix(eSide As eSideGrid, Settings As eSubTotalSettings, _
                     eRowInitial As Long, eColumnInitial As Long, _
                       eRowFinal As Long, eColumnFinal As Long) As Double
         bPrivateCellChange = True
         Dim nValue As Double
         nValue = 0
         Dim bFirst As Boolean
         bFirst = True
         Dim i As Long
         Dim j As Long
         Dim iDX As Integer
         
    If eSide = eLeftGrid Then
      iDX = 1
    Else
      iDX = 2
    End If

    With fg(iDX)
         For i = eRowInitial To eRowFinal
             For j = eColumnInitial To eColumnFinal
                 Select Case Settings
                        Case Is = axSTSum
                              nValue = nValue + CleanValue(.TextMatrix(i, j))
                        Case Is = axSTMax
                             If bFirst Then
                                nValue = CleanValue(.TextMatrix(i, j))
                                bFirst = False
                             Else
                                nValue = Max(nValue, CleanValue(.TextMatrix(i, j)))
                             End If
                        Case Is = axSTMin
                             If bFirst Then
                                nValue = CleanValue(.TextMatrix(i, j))
                                bFirst = False
                             Else
                                nValue = Min(nValue, CleanValue(.TextMatrix(i, j)))
                             End If
                        Case Is = axSTCount
                             nValue = nValue + IIf(Len(.TextMatrix(i, j)) > 0, 1, 0)
                 End Select
             Next j
         Next i
    End With
         bPrivateCellChange = False
        CalculateMatrix = nValue
End Property

Public Property Get CalculateRow(eSide As eSideGrid, Settings As eSubTotalSettings, eRow As Long, _
                                    eColInitial As Long, eColFinal As Long) As Double
         bPrivateCellChange = True
         Dim nValue As Double
         nValue = 0
         Dim bFirst As Boolean
         bFirst = True
         Dim i As Long
         Dim iDX As Integer
         
    If eSide = eLeftGrid Then
      iDX = 1
    Else
      iDX = 2
    End If
    
    With fg(iDX)
         For i = eColInitial To eColFinal
            Select Case Settings
                Case Is = axSTSum
                     nValue = nValue + CleanValue(.TextMatrix(eRow, i))
                Case Is = axSTMax
                     If bFirst Then
                        nValue = CleanValue(.TextMatrix(eRow, i))
                        bFirst = False
                     Else
                        nValue = Max(nValue, CleanValue(.TextMatrix(eRow, i)))
                     End If
                Case Is = axSTMin
                     If bFirst Then
                        nValue = CleanValue(.TextMatrix(eRow, i))
                        bFirst = False
                     Else
                        nValue = Min(nValue, CleanValue(.TextMatrix(eRow, i)))
                     End If
                Case Is = axSTCount
                        nValue = nValue + IIf(Len(.TextMatrix(eRow, i)) > 0, 1, 0)
            End Select
         Next i
         
         If Settings = axSTMultiply Then
            nValue = Val(.TextMatrix(eRow, eColInitial)) * CleanValue(.TextMatrix(eRow, eColFinal))
         End If
    End With
    
         bPrivateCellChange = False
        CalculateRow = nValue
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellAlignment() As AlignmentSettings
    CellAlignment = fg(1).CellAlignment
End Property

Public Property Let CellAlignment(ByVal New_CellAlignment As AlignmentSettings)
    fg(1).CellAlignment = New_CellAlignment
    fg(2).CellAlignment = New_CellAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellBackColor() As OLE_COLOR
    CellBackColor = txtInfoBar.BackColor
End Property

Public Property Let CellBackColor(ByVal New_CellBackColor As OLE_COLOR)
    fg(1).CellBackColor = New_CellBackColor
    fg(2).CellBackColor = New_CellBackColor
    'lblBar.BackColor = New_CellBackColor
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontBold() As Boolean
'    CellFontBold = fg(1).CellFontBold
'End Property
'
'Public Property Let CellFontBold(ByVal New_CellFontBold As Boolean)
'    fg(1).CellFontBold = New_CellFontBold
'    fg(2).CellFontBold = New_CellFontBold
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontItalic() As Boolean
'    CellFontItalic = fg(1).CellFontItalic
'End Property
'
'Public Property Let CellFontItalic(ByVal New_CellFontItalic As Boolean)
'    fg(1).CellFontItalic = New_CellFontItalic
'    fg(2).CellFontItalic = New_CellFontItalic
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontName() As String
'    CellFontName = fg(1).CellFontName
'End Property
'
'Public Property Let CellFontName(ByVal New_CellFontName As String)
'    fg(1).CellFontName = New_CellFontName
'    fg(2).CellFontName = New_CellFontName
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontSize() As Single
'    CellFontSize = fg(1).CellFontSize
'End Property
'
'Public Property Let CellFontSize(ByVal New_CellFontSize As Single)
'    fg(1).CellFontSize = New_CellFontSize
'    fg(2).CellFontSize = New_CellFontSize
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontStrikeThrough() As Boolean
'    CellFontStrikeThrough = fg(1).CellFontStrikeThrough
'End Property
'
'Public Property Let CellFontStrikeThrough(ByVal New_CellFontStrikeThrough As Boolean)
'    fg(1).CellFontStrikeThrough = New_CellFontStrikeThrough
'    fg(2).CellFontStrikeThrough = New_CellFontStrikeThrough
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontUnderline() As Boolean
'    CellFontUnderline = fg(1).CellFontUnderline
'End Property
'
'Public Property Let CellFontUnderline(ByVal New_CellFontUnderline As Boolean)
'    fg(1).CellFontUnderline = New_CellFontUnderline
'    fg(2).CellFontUnderline = New_CellFontUnderline
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get CellFontWidth() As Single
'    CellFontWidth = fg(1).CellFontWidth
'End Property
'
'Public Property Let CellFontWidth(ByVal New_CellFontWidth As Single)
'    fg(1).CellFontWidth = New_CellFontWidth
'    fg(2).CellFontWidth = New_CellFontWidth
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellForeColor() As OLE_COLOR
    CellForeColor = fg(1).CellForeColor
End Property

Public Property Let CellForeColor(ByVal New_CellForeColor As OLE_COLOR)
    fg(1).CellForeColor = New_CellForeColor
    fg(2).CellForeColor = New_CellForeColor
End Property

Public Property Get cell(Setting As eCellProperty, ByVal Row1 As Long, ByVal Col1 As Long, _
                         ByVal Row2 As Long, ByVal Col2 As Long, eG As eSideGrid) As Variant
    bPrivateCellChange = True
    Dim OldRow As Long
    Dim OldCol As Long
    
With fg(eG)
    OldRow = .Row
    OldCol = .Col
    .Row = Row1
    .Col = Col1
    Select Case Setting
        Case Is = axcpCellAlignment
            cell = .CellAlignment
        Case Is = axcpCellFontBold
            cell = .CellFontBold
        Case Is = axcpCellFontName
            cell = .CellFontName
        Case Is = axcpCellFontSize
            cell = .CellFontSize
        Case Is = axcpCellForeColor
            cell = .CellForeColor
        Case Is = axcpCellBackColor
            cell = .CellBackColor
    End Select
    .Row = OldRow
    .Col = OldCol
End With
    bPrivateCellChange = False
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellHeight
Public Property Get CellHeight(eG As eSideGrid) As Long
    CellHeight = fg(eG).CellHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellLeft
Public Property Get CellLeft(eG As eSideGrid) As Long
    CellLeft = fg(eG).CellLeft
End Property

Public Property Let cell(Setting As eCellProperty, ByVal Row1 As Long, ByVal Col1 As Long, _
                         ByVal Row2 As Long, ByVal Col2 As Long, eG As eSideGrid, New_Val As Variant)
    bPrivateCellChange = True
    Dim OldRow As Long
    Dim OldCol As Long

With fg(eG)
    OldRow = .Row
    OldCol = .Col
    Dim i As Long
    Dim j As Long
    For i = Row1 To Row2
        For j = Col1 To Col2
            .Row = i
            .Col = j
            Select Case Setting
                Case Is = axcpCellAlignment
                    .CellAlignment = New_Val
                Case Is = axcpCellFontBold
                    .CellFontBold = New_Val
                Case Is = axcpCellFontName
                    .CellFontName = New_Val
                Case Is = axcpCellFontSize
                    .CellFontSize = New_Val
                Case Is = axcpCellForeColor
                    .CellForeColor = New_Val
                Case Is = axcpCellBackColor
                    .CellBackColor = New_Val
            End Select
        Next j
    Next i
    .Row = OldRow
    .Col = OldCol
End With

    bPrivateCellChange = False
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellPictureAlignment(eG As eSideGrid) As AlignmentSettings
    CellPictureAlignment = fg(eG).CellPictureAlignment
End Property

Public Property Let CellPictureAlignment(eG As eSideGrid, ByVal New_CellPictureAlignment As AlignmentSettings)
    fg(eG).CellPictureAlignment = New_CellPictureAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellPicture
Public Property Get CellPicture(eG As eSideGrid) As Picture
    Set CellPicture = fg(eG).CellPicture
End Property

Public Property Set CellPicture(eG As eSideGrid, ByVal New_CellPicture As Picture)
    Set fg(eG).CellPicture = New_CellPicture
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellTextStyle(eG As eSideGrid) As TextStyleSettings
    CellTextStyle = fg(eG).CellTextStyle
End Property

Public Property Let CellTextStyle(eG As eSideGrid, ByVal New_CellTextStyle As TextStyleSettings)
    fg(eG).CellTextStyle = New_CellTextStyle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellTop
Public Property Get CellTop(eG As eSideGrid) As Long
    CellTop = fg(eG).CellTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellWidth
Public Property Get CellWidth(eG As eSideGrid) As Long
    CellWidth = fg(eG).CellWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Clip(eG As eSideGrid) As String
    Clip = fg(eG).Clip
End Property

Public Property Let Clip(eG As eSideGrid, ByVal New_Clip As String)
    fg(eG).Clip = New_Clip
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColAlignment
Public Property Get ColAlignment(eG As eSideGrid, ByVal Index As Long) As Integer
    ColAlignment = fg(eG).ColAlignment(Index)
End Property

Public Property Let ColAlignment(eG As eSideGrid, ByVal Index As Long, ByVal New_ColAlignment As Integer)
    fg(eG).ColAlignment(Index) = New_ColAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColData
Public Property Get ColData(eG As eSideGrid, ByVal Index As Long) As Long
    ColData = fg(eG).ColData(Index)
End Property

Public Property Let ColData(eG As eSideGrid, ByVal Index As Long, ByVal New_ColData As Long)
    fg(eG).ColData(Index) = New_ColData
End Property

Public Property Get ColDisplayFormat(ByVal Col As Long) As String
    ColDisplayFormat = m_ColsR(Col).ColDisplayFormat
End Property

Public Property Let ColDisplayFormat(ByVal Col As Long, ByVal New_ColDisplayFormat As String)
    m_ColsR(Col).ColDisplayFormat = New_ColDisplayFormat
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Col(eG As eSideGrid) As Long
    Col = fg(eG).Col
End Property

Public Property Get ColInputMask(ByVal Col As Long) As String
    ColInputMask = m_ColsR(Col).ColInputMask
End Property

Public Property Let ColInputMask(ByVal Col As Long, ByVal New_ColInputMask As String)
    m_ColsR(Col).ColInputMask = New_ColInputMask
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColIsVisible
Public Property Get ColIsVisible(eG As eSideGrid, ByVal Index As Long) As Boolean
    ColIsVisible = fg(eG).ColIsVisible(Index)
End Property

Public Property Let Col(eG As eSideGrid, ByVal New_Col As Long)
    fg(eG).Col = New_Col
End Property

'''''''''''''TESTING''''''''''''''''''
Public Property Get ColObject(eG As eSideGrid, ByVal Col As Long) As eColumnType
If eG = eLeftGrid Then
    ColObject = m_ColsL(Col).ColType
ElseIf eG = eRightGrid Then
    ColObject = m_ColsR(Col).ColType
End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColPos
Public Property Get ColPos(eG As eSideGrid, ByVal Index As Long) As Long
    ColPos = fg(eG).ColPos(Index)
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=fg,fg,-1,ColPosition
'Public Property Get ColPosition(ByVal index As Long) As Long
'    ColPosition = fg.ColPosition(index)
'End Property

Public Property Let ColPosition(eG As eSideGrid, ByVal Index As Long, ByVal New_ColPosition As Long)
    fg(eG).ColPosition(Index) = New_ColPosition
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ColSel(eG As eSideGrid) As Long
    ColSel = fg(eG).ColSel
End Property

Public Property Let ColSel(eG As eSideGrid, ByVal New_ColSel As Long)
    fg(eG).ColSel = New_ColSel
End Property

Public Property Get ColsLeft() As Long
    ColsLeft = fg(1).Cols
End Property

Public Property Let ColsLeft(ByVal New_Cols As Long)
    On Error GoTo ErrHand
    fg(1).Cols() = New_Cols
    ReDim Preserve m_ColsL(New_Cols + 1)
    PropertyChanged "ColsLeft"
    Exit Property
ErrHand:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Public Property Get ColsRight() As Long
    ColsRight = fg(2).Cols
End Property

Public Property Let ColsRight(ByVal New_Cols As Long)
    On Error GoTo ErrHand
    fg(2).Cols() = New_Cols
    ReDim Preserve m_ColsR(New_Cols + 1)
    PropertyChanged "ColsRight"
    Exit Property
ErrHand:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property


'***************
' New Properties
'***************

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColWidth
Public Property Get ColWidth(eG As eSideGrid2, ByVal Index As Long) As Long
Select Case eG
  Case Is = LeftGrid
        ColWidth = fg(1).ColWidth(Index)
  Case Is = RightGrid
        ColWidth = fg(2).ColWidth(Index)
  Case Is = BothGrids
        ColWidth = fg(1).ColWidth(Index)
End Select
End Property

Public Property Let ColWidth(eG As eSideGrid2, ByVal Index As Long, ByVal New_ColWidth As Long)
Select Case eG
  Case Is = LeftGrid
    fg(1).ColWidth(Index) = New_ColWidth
  Case Is = RightGrid
    fg(2).ColWidth(Index) = New_ColWidth
  Case Is = BothGrids
    fg(1).ColWidth(Index) = New_ColWidth
    fg(2).ColWidth(Index) = New_ColWidth
End Select
End Property

Public Property Get Editable() As Boolean
    Editable = m_Editable
End Property

Public Property Let Editable(ByVal New_Editable As Boolean)
    m_Editable = New_Editable
    PropertyChanged "Editable"
End Property

Public Property Get EditSelLength() As Long
    EditSelLength = txtEdit.SelLength
End Property

Public Property Let EditSelLength(ByVal NewData As Long)
    txtEdit.SelLength = NewData
End Property

Public Property Get EditSelStart() As Long
    EditSelStart = txtEdit.SelStart
End Property

Public Property Let EditSelStart(ByVal NewData As Long)
    txtEdit.SelStart = NewData
End Property

Public Property Get EditSelText() As String
    EditSelText = txtEdit.SelText
End Property

Public Property Let EditSelText(ByVal NewData As String)
    txtEdit.SelText = NewData
End Property

Public Property Get EditText() As String
   EditText = txtEdit.Text
End Property

Public Property Let EditText(ByVal NewData As String)
   txtEdit.Text = NewData
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
    'Enabled = fg(1).Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
UserControl.Enabled = New_Enabled
'    fg(1).Enabled() = New_Enabled
'    fg(2).Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get EnterKeyBehaviour() As eEnterkeyBehaviour
    EnterKeyBehaviour = m_EnterKeyBehaviour
End Property

Public Property Let EnterKeyBehaviour(ByVal New_EnterKeyBehaviour As eEnterkeyBehaviour)
    m_EnterKeyBehaviour = New_EnterKeyBehaviour
    PropertyChanged "EnterKeyBehaviour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FillStyle
Public Property Get FillStyle(eG As eSideGrid) As FillStyleSettings
    FillStyle = fg(eG).FillStyle
End Property

Public Property Let FillStyle(eG As eSideGrid, ByVal New_FillStyle As FillStyleSettings)
    fg(eG).FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FixedAlignment
Public Property Get FixedAlignment(eG As eSideGrid, ByVal Index As Long) As Integer
    FixedAlignment = fg(eG).FixedAlignment(Index)
End Property

Public Property Let FixedAlignment(eG As eSideGrid, ByVal Index As Long, ByVal New_FixedAlignment As Integer)
    fg(eG).FixedAlignment(Index) = New_FixedAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FixedCols
Public Property Get FixedCols() As Long
    FixedCols = fg(1).FixedCols
End Property

Public Property Let FixedCols(ByVal New_FixedCols As Long)
    fg(1).FixedCols() = New_FixedCols
    PropertyChanged "FixedCols"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FixedRows
Public Property Get FixedRows() As Long
    FixedRows = fg(1).FixedRows
End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)
    fg(1).FixedRows() = New_FixedRows
    fg(2).FixedRows() = New_FixedRows
    PropertyChanged "FixedRows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FocusRect
Public Property Get FocusRect() As FocusRectSettings
    FocusRect = fg(1).FocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As FocusRectSettings)
    fg(1).FocusRect() = New_FocusRect
    fg(2).FocusRect() = New_FocusRect
    PropertyChanged "FocusRect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontBold(eG As eSideGrid) As Boolean
    FontBold = fg(eG).FontBold
End Property

Public Property Let FontBold(eG As eSideGrid, ByVal New_FontBold As Boolean)
    fg(eG).FontBold = New_FontBold
    txtEdit.FontBold = New_FontBold
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Font
Public Property Get Font() As Font
    Set Font = fg(1).Font
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get FontItalic(eG As eSideGrid) As Boolean
    FontItalic = fg(eG).FontItalic
End Property

Public Property Let FontItalic(eG As eSideGrid, ByVal New_FontItalic As Boolean)
    fg(eG).FontItalic = New_FontItalic
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontName(eG As eSideGrid) As String
    FontName = fg(eG).FontName
End Property

Public Property Let FontName(eG As eSideGrid, ByVal New_FontName As String)
    fg(eG).FontName = New_FontName
    txtEdit.FontName = New_FontName
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set fg(1).Font = New_Font
    Set fg(2).Font = New_Font
    Set txtEdit.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontSize(eG As eSideGrid) As Long
    FontSize = fg(eG).FontSize
End Property

Public Property Let FontSize(eG As eSideGrid, ByVal New_FontSize As Long)
    fg(eG).FontSize = New_FontSize
    txtEdit.FontSize = New_FontSize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontStrikethru(eG As eSideGrid) As Boolean
    FontStrikethru = fg(eG).FontStrikethru
End Property

Public Property Let FontStrikethru(eG As eSideGrid, ByVal New_FontStrikethru As Boolean)
    fg(eG).FontStrikethru = New_FontStrikethru
    txtEdit.FontStrikethru = New_FontStrikethru
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontUnderline(eG As eSideGrid) As Boolean
    FontUnderline = fg(eG).FontUnderline
End Property

Public Property Let FontUnderline(eG As eSideGrid, ByVal New_FontUnderline As Boolean)
    fg(eG).FontUnderline = New_FontUnderline
    txtEdit.FontUnderline = New_FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FontWidth
Public Property Get FontWidth(eG As eSideGrid) As Single
    FontWidth = fg(eG).FontWidth
End Property

Public Property Let FontWidth(eG As eSideGrid, ByVal New_FontWidth As Single)
    fg(eG).FontWidth() = New_FontWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ForeColorFixed
Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = fg(1).ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    fg(1).ForeColorFixed() = fg(2).ForeColorFixed() = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = fg(1).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    fg(1).ForeColor() = fg(1).ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ForeColorSel
Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = fg(1).ForeColorSel
End Property

Public Property Let ForeColorSel(ByVal New_ForeColorSel As OLE_COLOR)
    fg(1).ForeColorSel() = fg(2).ForeColorSel() = New_ForeColorSel
    PropertyChanged "ForeColorSel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FormatString
Public Property Get FormatString(eG As eSideGrid) As String
    FormatString = fg(eG).FormatString
End Property

Public Property Let FormatString(eG As eSideGrid, ByVal New_FormatString As String)
    fg(eG).FormatString() = New_FormatString
    PropertyChanged "FormatString"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,GridColorFixed
Public Property Get GridColorFixed() As OLE_COLOR
    GridColorFixed = fg(1).GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
    fg(1).GridColorFixed() = fg(2).GridColorFixed() = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,GridColor
Public Property Get GridColor() As OLE_COLOR
    GridColor = fg(1).GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    fg(1).GridColor() = fg(2).GridColor() = New_GridColor
    PropertyChanged "GridColor"
End Property

Public Property Get GridLinesFixed() As GridLineSettings
    GridLinesFixed = fg(1).GridLineWidth
End Property

Public Property Let GridLinesFixed(ByVal New_GridLinesFixed As GridLineSettings)
    fg(1).GridLinesFixed() = New_GridLinesFixed
    fg(2).GridLinesFixed() = New_GridLinesFixed
    PropertyChanged "GridLinesFixed"
End Property

Public Property Get GridLines() As GridLineSettings
   GridLines = fg(1).GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As GridLineSettings)
  fg(1).GridLines = New_GridLines
  fg(1).GridLines = New_GridLines
  PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,GridLineWidth
Public Property Get GridLineWidth() As Integer
    GridLineWidth = fg(1).GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Integer)
    fg(1).GridLineWidth() = New_GridLineWidth
    fg(2).GridLineWidth() = New_GridLineWidth
    PropertyChanged "GridLineWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,HighLight
Public Property Get HighLight() As HighLightSettings
    HighLight = fg(1).HighLight
End Property

Public Property Let HighLight(ByVal New_HighLight As HighLightSettings)
    fg(1).HighLight() = New_HighLight
    fg(2).HighLight() = New_HighLight
    PropertyChanged "HighLight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get hwnd(eG As eSideGrid) As Long
    hwnd = fg(eG).hwnd
End Property

Public Property Get InfoBarText() As String
  InfoBarText = txtInfoBar.Text
End Property

Public Property Let InfoBarText(sText As String)
   m_txtInfoBar = sText
   txtInfoBar.Text = m_txtInfoBar
   PropertyChanged "InfoBarText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get LeftCol(eG As eSideGrid) As Long
    LeftCol = fg(eG).LeftCol
End Property

Public Property Let LeftCol(eG As eSideGrid, ByVal New_LeftCol As Long)
    fg(eG).LeftCol = New_LeftCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MergeCells
Public Property Get MergeCells(eG As eSideGrid) As MergeCellsSettings
    MergeCells = fg(eG).MergeCells
End Property

Public Property Let MergeCells(eG As eSideGrid, ByVal New_MergeCells As MergeCellsSettings)
    fg(eG).MergeCells() = New_MergeCells
    PropertyChanged "MergeCells"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MergeCol
Public Property Get MergeCol(eG As eSideGrid, ByVal Index As Long) As Boolean
    MergeCol = fg(eG).MergeCol(Index)
End Property

Public Property Let MergeCol(eG As eSideGrid, ByVal Index As Long, ByVal New_MergeCol As Boolean)
    fg(eG).MergeCol(Index) = New_MergeCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MergeRow
Public Property Get MergeRow(eG As eSideGrid, ByVal Index As Long) As Boolean
    MergeRow = fg(eG).MergeRow(Index)
End Property

Public Property Let MergeRow(eG As eSideGrid, ByVal Index As Long, ByVal New_MergeRow As Boolean)
    fg(eG).MergeRow(Index) = New_MergeRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MouseCol
Public Property Get MouseCol() As Long
    MouseCol = fg(2).MouseCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = fg(1).MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set fg(1).MouseIcon = New_MouseIcon
    Set fg(2).MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MousePointer
Public Property Get MousePointer(eG As eSideGrid) As MousePointerSettings
    MousePointer = fg(eG).MousePointer
End Property

Public Property Let MousePointer(eG As eSideGrid, ByVal New_MousePointer As MousePointerSettings)
    fg(eG).MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MouseRow
Public Property Get MouseRow(eG As eSideGrid) As Long
    MouseRow = fg(eG).MouseRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,OLEDropMode
Public Property Get OLEDropMode(eG As eSideGrid) As OLEDropConstants
    OLEDropMode = fg(eG).OLEDropMode
End Property

Public Property Let OLEDropMode(eG As eSideGrid, ByVal New_OLEDropMode As OLEDropConstants)
    fg(eG).OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Picture
Public Property Get Picture(eG As eSideGrid) As Picture
    Set Picture = fg(eG).Picture
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,PictureType
Public Property Get PictureType(eG As eSideGrid) As PictureTypeSettings
    PictureType = fg(eG).PictureType
End Property

Public Property Let PictureType(eG As eSideGrid, ByVal New_PictureType As PictureTypeSettings)
    fg(eG).PictureType() = New_PictureType
    PropertyChanged "PictureType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Redraw
Public Property Get Redraw() As Boolean
    Redraw = fg(1).Redraw
End Property

Public Property Let Redraw(ByVal New_Redraw As Boolean)
    fg(1).Redraw() = New_Redraw
    fg(2).Redraw() = New_Redraw
    PropertyChanged "Redraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RightToLeft
Public Property Get RightToLeft(eG As eSideGrid) As Boolean
    RightToLeft = fg(eG).RightToLeft
End Property

Public Property Let RightToLeft(eG As eSideGrid, ByVal New_RightToLeft As Boolean)
    fg(eG).RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowData
Public Property Get RowData(ByVal Index As Long) As Long
    RowData = fg(2).RowData(Index)
End Property

Public Property Let RowData(ByVal Index As Long, ByVal New_RowData As Long)
    fg(2).RowData(Index) = New_RowData
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Row(eG As eSideGrid) As Long
    Row = fg(eG).Row
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowHeight
Public Property Get RowHeight(ByVal Index As Long) As Long
    RowHeight = fg(1).RowHeight(Index)
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal New_RowHeight As Long)
    fg(1).RowHeight(Index) = New_RowHeight
    fg(2).RowHeight(Index) = New_RowHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowHeightMin
Public Property Get RowHeightMin() As Long
    RowHeightMin = fg(1).RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    fg(1).RowHeightMin() = New_RowHeightMin
    fg(2).RowHeightMin() = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowIsVisible
Public Property Get RowIsVisible(ByVal Index As Long) As Boolean
    RowIsVisible = fg(1).RowIsVisible(Index)
End Property

Public Property Let Row(eG As eSideGrid, ByVal New_Row As Long)
    fg(eG).Row = New_Row
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowPos
Public Property Get RowPos(ByVal Index As Long) As Long
    RowPos = fg(1).RowPos(Index)
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=fg,fg,-1,RowPosition
'Public Property Get RowPosition(ByVal index As Long) As Long
'    RowPosition = fg(1).RowPosition(index)
'End Property

Public Property Let RowPosition(ByVal Index As Long, ByVal New_RowPosition As Long)
    fg(1).RowPosition(Index) = New_RowPosition
    fg(2).RowPosition(Index) = New_RowPosition
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get RowSel() As Long
    RowSel = fg(1).RowSel
End Property

Public Property Let RowSel(ByVal New_RowSel As Long)
    fg(1).RowSel = New_RowSel
    fg(2).RowSel = New_RowSel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Rows
Public Property Get Rows() As Long
    Rows = fg(1).Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
    Dim mOldRows As Long
    mOldRows = fg(1).Rows
    
    fg(1).Rows() = New_Rows
    fg(2).Rows() = New_Rows
    
    If New_Rows > mOldRows Then
        If m_BackColorAlternate <> fg(1).BackColor Then
            Call SetAlternateRowColors(1, m_BackColorAlternate, fg(1).BackColor)
            Call SetAlternateRowColors(2, m_BackColorAlternate, fg(1).BackColor)
        End If
    End If
    
    PropertyChanged "Rows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ScrollBars
Public Property Get ScrollBars() As ScrollBarsSettings
    ScrollBars = fg(2).ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
    fg(1).ScrollBars() = New_ScrollBars
    fg(2).ScrollBars() = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ScrollTrack
Public Property Get ScrollTrack() As Boolean
    ScrollTrack = fg(1).ScrollTrack
End Property

Public Property Let ScrollTrack(ByVal New_ScrollTrack As Boolean)
    fg(1).ScrollTrack() = New_ScrollTrack
    fg(2).ScrollTrack() = New_ScrollTrack
    PropertyChanged "ScrollTrack"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,SelectionMode
Public Property Get SelectionMode() As SelectionModeSettings
    SelectionMode = fg(1).SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
    fg(1).SelectionMode() = New_SelectionMode
    fg(2).SelectionMode() = New_SelectionMode
    PropertyChanged "SelectionMode"
End Property

Public Property Let SetColObject(eG As eSideGrid, ByVal Col As Long, SetVisibleOnCellTextChange As Boolean, ByVal new_ColType As eColumnType)
If eG = eLeftGrid Then
    m_ColsL(Col).ColType = new_ColType
ElseIf eG = eRightGrid Then
    m_ColsR(Col).ColType = new_ColType
End If
If SetVisibleOnCellTextChange Then SetColumnObject CInt(eG), Col
End Property

Public Property Get SetInfoBar() As eTypeInfoBar
    SetInfoBar = m_SetInfoBar
End Property

Public Property Let SetInfoBar(eNewType As eTypeInfoBar)
  m_SetInfoBar = eNewType
  PropertyChanged "SetInfoBar"
End Property

Public Property Get ShowInfoBar() As Boolean
    ShowInfoBar = txtInfoBar.Visible
End Property

Public Property Let ShowInfoBar(ByVal bShowBar As Boolean)
      txtInfoBar.Visible = bShowBar
      UserControl_Resize
      txtInfoBar.Appearance = Appearance
      PropertyChanged "ShowInfoBar"
End Property

Public Property Get SortColumnMode() As SortSettings
    SortColumnMode = m_SortMode
End Property

Public Property Let SortColumnMode(ByVal New_SortMode As SortSettings)
    m_SortMode = New_SortMode
    PropertyChanged "SortColumnMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Sort
Public Property Let Sort(eG As eSideGrid, ByVal New_Sort As SortSettings)
    fg(eG).Sort() = New_Sort
End Property

Public Property Get SplitterFixed() As Boolean
    SplitterFixed = m_SplitterFixed
End Property

Public Property Let SplitterFixed(ByVal N_SplitterFixed As Boolean)
    Divisor.Enabled = Not N_SplitterFixed
    m_SplitterFixed = N_SplitterFixed
    PropertyChanged "SplitterFixed"
End Property

Public Property Get SplitterPos() As Long
    SplitterPos = m_SplitterPos
End Property

Public Property Let SplitterPos(ByVal N_SplitterPos As Long)
    m_SplitterPos = N_SplitterPos
    PropertyChanged "SplitterPos"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextArray
Public Property Get TextArray(eG As eSideGrid, ByVal Index As Long) As String
    TextArray = fg(eG).TextArray(Index)
End Property

Public Property Let TextArray(eG As eSideGrid, ByVal Index As Long, ByVal New_TextArray As String)
    fg(eG).TextArray(Index) = New_TextArray
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Text(eG As eSideGrid) As String
    Text = fg(eG).Text
End Property

Public Property Let Text(eG As eSideGrid, ByVal New_Text As String)
    Call DisplayFormatedText(eG, New_Text, fg(eG).Row, fg(eG).Col)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextMatrix
Public Property Get TextMatrix(eG As eSideGrid, ByVal Row As Long, ByVal Col As Long) As String
    TextMatrix = fg(eG).TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(eG As eSideGrid, ByVal Row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
    Call DisplayFormatedText(eG, New_TextMatrix, Row, Col)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextStyleFixed
Public Property Get TextStyleFixed(eG As eSideGrid) As TextStyleSettings
    TextStyleFixed = fg(eG).TextStyleFixed
End Property

Public Property Let TextStyleFixed(eG As eSideGrid, ByVal New_TextStyleFixed As TextStyleSettings)
    fg(eG).TextStyleFixed() = New_TextStyleFixed
    PropertyChanged "TextStyleFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextStyle
Public Property Get TextStyle(eG As eSideGrid) As TextStyleSettings
    TextStyle = fg(eG).TextStyle
End Property

Public Property Let TextStyle(eG As eSideGrid, ByVal New_TextStyle As TextStyleSettings)
    fg(eG).TextStyle() = New_TextStyle
    PropertyChanged "TextStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get TopRow(eG As eSideGrid) As Long
    TopRow = fg(eG).TopRow
End Property

Public Property Let TopRow(eG As eSideGrid, ByVal New_TopRow As Long)
    fg(eG).TopRow = New_TopRow
End Property

Public Property Get Value(eG As eSideGrid) As Variant
     Value = Val(fg(eG).Text)
End Property

Public Property Get ValueMatrix(eG As eSideGrid, ByVal Row As Long, ByVal Col As Long) As Double
     ValueMatrix = CleanValue(fg(eG).TextMatrix(Row, Col))
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,WordWrap
Public Property Get WordWrap(eG As eSideGrid) As Boolean
    WordWrap = fg(eG).WordWrap
End Property

Public Property Let WordWrap(eG As eSideGrid, ByVal New_WordWrap As Boolean)
    fg(eG).WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property










