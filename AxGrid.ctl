VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl AxGrid 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   PropertyPages   =   "AxGrid.ctx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   4935
   ToolboxBitmap   =   "AxGrid.ctx":003F
   Begin VB.PictureBox pUncheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3180
      Picture         =   "AxGrid.ctx":0571
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox pCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2910
      Picture         =   "AxGrid.ctx":0883
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   435
      TabIndex        =   1
      Top             =   2610
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txtInfoBar 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3585
      Visible         =   0   'False
      Width           =   5865
   End
   Begin AxioGrid.axComboBox cBox 
      Height          =   255
      Left            =   2670
      TabIndex        =   5
      Top             =   2655
      Visible         =   0   'False
      Width           =   690
      _ExtentX        =   1217
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
      ItemSelectMode  =   1
      ListIndex       =   -1
      ItemSelectMode  =   1
   End
   Begin VB.TextBox TBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFF4&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1365
      TabIndex        =   4
      Top             =   2715
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Boton 
      Caption         =   "<"
      Height          =   255
      Left            =   3615
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.ListBox LBox 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   3960
      TabIndex        =   2
      Top             =   2445
      Visible         =   0   'False
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid fg 
      Height          =   3315
      Left            =   255
      TabIndex        =   0
      Top             =   135
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      BackColorFixed  =   -2147483626
      GridColorFixed  =   8421504
      Appearance      =   0
   End
   Begin VB.Menu MnuFGridRows 
      Caption         =   "Filas"
      Begin VB.Menu MnuFGridAddRow 
         Caption         =   "Insertar Fila"
      End
      Begin VB.Menu mnuDeleteGridRow 
         Caption         =   "Borrar Fila"
      End
      Begin VB.Menu MnuFGridAddRowEnd 
         Caption         =   "Añadir Fila al final"
      End
   End
   Begin VB.Menu MnuFGridCols 
      Caption         =   "Columnas"
      Begin VB.Menu MnuFGridAddCol 
         Caption         =   "Insertar Columna"
      End
      Begin VB.Menu mnuDeleteGridCol 
         Caption         =   "Borrar Columna"
      End
      Begin VB.Menu MnuFGridAddColEnd 
         Caption         =   "Añadir Columna al final"
      End
   End
End
Attribute VB_Name = "AxGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'''-----------------------------------------------------------
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'''-----------------------------------------------------------

'Extra Events
Public Event BeforeEdit(Cancel As Boolean)
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
Public Event KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Public Event ValidateEdit(Row As Long, Col As Long, Cancel As Boolean)
Public Event ButtonClick(ByVal Row As Long, ByVal Col As Long)
Public Event CellTextChange(ByVal Row As Long, ByVal Col As Long)
Public Event ListClick(ByVal Row As Long, ByVal Col As Long, ByVal iListIndex As Long)

'Mapped Events
Public Event Click(lRow As Long, lCol As Long)
Public Event Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
Public Event DblClick(lRow As Long, lCol As Long)
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
Public Event OLEStartDrag(Data As MSFlexGridLib.DataObject, AllowedEffects As Long)
Public Event OLESetData(Data As MSFlexGridLib.DataObject, DataFormat As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLEDragOver(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)

Private m_Cols()  As FgCol
Private cFlatSb   As New cCoolScrollbars
'Private cFlat     As New cFlatControl

Dim bPrivateCellChange    As Boolean
Dim m_EnterKeyBehaviour   As eEnterkeyBehaviour
Dim m_AutoSizeMode        As eAutoSizeSetting
Dim m_BackColorAlternate  As OLE_COLOR
Dim m_SortMode            As SortSettings
Dim m_ColType             As eColumnType
Dim m_ColObject           As Long
Dim m_Editable            As Boolean
Dim m_Command             As String
Dim m_SetInfoBar          As eTypeSingleInfoBar
Dim m_txtInfoBar          As String
Dim m_ShowMenuRow         As Boolean
Dim m_ShowMenuCol         As Boolean
Dim m_AutoNumCol          As Boolean
Dim oRow                  As Long
Dim oCol                  As Long
Dim lCol                  As Long
Dim lRow                  As Long
Dim bCboxSel              As Boolean
Dim sTextOnCell           As String

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
      .lblProd.Caption = "AxGrid "
      .lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision & " "
      .lblV1.Caption = App.Major
      .lblV2.Caption = "." & App.Minor
      .lblInfo.Caption = "Another Xtended Grid."
      .Show (vbModal)
  End With
End Sub

Public Sub AddItem(strValue As String, Optional rowIndex As Long)
If Not IsMissing(rowIndex) Then
  fg.AddItem strValue, rowIndex
Else
  fg.AddItem strValue
End If

End Sub

Public Sub AddItemObject(strValue As String, Optional oIndex As Long)
If Not IsMissing(oIndex) Then
  cBox.AddItem strValue, oIndex
  LBox.AddItem strValue, oIndex
Else
  cBox.AddItem strValue
  LBox.AddItem strValue
End If
End Sub

Public Sub ClearItemObject(cObject As eTypeControl)
  If cObject = oListBox Then LBox.Clear
  If cObject = oComboBox Then cBox.Clear
End Sub

Public Sub AutoSizeCols(eFirstCol As Long, eLastCol As Long)
 Dim i As Long, j As Long
 Dim nMaxWidth As Long
 Dim nCurrWidth As Long
                If IsMissing(eLastCol) Then eLastCol = eFirstCol
                Call AutoSizeC(fg, eFirstCol, eLastCol, True)
End Sub

Public Sub ClearGrid()
  fg.Clear
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

Public Sub ExtendLastColumn(Optional Col As Long)
    Dim m_lScrollWidth As Long
    m_lScrollWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelY
    
    Dim lTotWidth As Long
    Dim lScrollWidth As Long
    Dim nMargin As Long
    nMargin = 95
    With fg
        ' is there a vertical scrollbar
        lScrollWidth = 0
        If .ScrollBars = flexScrollBarBoth Or .ScrollBars = flexScrollBarBoth Then
            If Not .RowIsVisible(0) Or Not .RowIsVisible(.Rows - 1) Then
                lScrollWidth = m_lScrollWidth
            End If
        End If
    
        Dim nWidth As Long
        nWidth = fg.Width - lScrollWidth
        Dim nColWidths As Long
        Dim i As Long
        For i = 0 To fg.Cols - 1
            nColWidths = nColWidths + fg.ColWidth(i) + fg.GridLineWidth
        Next
        If fg.Appearance = flex3D Then
            nMargin = 95
        Else
            nMargin = 35
        End If
        
        Dim nProcessCol As Long
        If Not IsMissing(Col) Then
           nProcessCol = Col
        Else
           nProcessCol = fg.Cols - 1
        End If
        If nColWidths < nWidth - nMargin Then
            fg.ColWidth(nProcessCol) = fg.ColWidth(nProcessCol) + (nWidth - nColWidths - nMargin)
        End If
        If lScrollWidth = 0 Then
        End If
    End With
End Sub

Public Function GridToHTML() As String
    Dim i As Long
    Dim j As Long
    Dim sText As String
    sText = "<HTML>" & vbCrLf
    sText = sText & "<BODY>" & vbCrLf
    sText = sText & "<TABLE>" & vbCrLf
    For i = 0 To fg.Rows - 1
       sText = sText & "<TR>" & vbCrLf
       For j = 0 To fg.Cols - 1
           sText = sText & "<TD>" & fg.TextMatrix(i, j) & "</TD>"
       Next
       sText = sText & vbCrLf & "</TR>" & vbCrLf
    Next
    sText = sText & "</TABLE>" & vbCrLf
    sText = sText & "</BODY>" & vbCrLf
    sText = sText & "</HTML>"
    
    GridToHTML = sText

ErrHand:
    Err.Raise Err.Number, Err.Source, Err.Description
    GridToHTML = ""
End Function

Public Sub RemoveItem(Index As Long)
  fg.RemoveItem (Index)
End Sub

Public Sub ResetGrid()
Dim R As Long
  With fg
    For R = 0 To .Cols - 1
        fg.RemoveItem (R)
    Next R
    .Clear
  End With
End Sub

Public Sub SaveAsHTML(FileName As String)
    Dim i As Long
    Dim j As Long
    Dim sText As String
    sText = "<HTML>" & vbCrLf
    sText = sText & "<BODY>" & vbCrLf
    sText = sText & "<TABLE>" & vbCrLf
    For i = 0 To fg.Rows - 1
       sText = sText & "<TR>" & vbCrLf
       For j = 0 To fg.Cols - 1
           sText = sText & "<TD>" & fg.TextMatrix(i, j) & "</TD>"
       Next
       sText = sText & vbCrLf & "</TR>" & vbCrLf
    Next
    sText = sText & "</TABLE>" & vbCrLf
    sText = sText & "</BODY>" & vbCrLf
    sText = sText & "</HTML>"
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

Public Function SaveAsOldXLS(sFilename As String)
    Dim myExcel As ExcelFileV2
    Dim excelDouble As Double
    Dim rowOffset As Long
    Dim aTemp() As String
      
    Set myExcel = New ExcelFileV2
      
      With myExcel
          .OpenFile sFilename
          ' FlexGrid -> Fixedrows
        For lRow = 1 To fg.FixedRows
          For lCol = 1 To fg.Cols
            .EWriteString lRow + rowOffset, lCol, fg.TextMatrix(lRow - 1, lCol - 1)
          Next lCol
        Next lRow
      
      ' Data
        For lRow = fg.FixedRows + 1 To fg.Rows
          ' FlexGrid -> Fixedcols
          For lCol = 1 To fg.FixedCols
            .EWriteString lRow + rowOffset, lCol, fg.TextMatrix(lRow - 1, lCol - 1)
          Next lCol
          
          ' FlexGrid -> Data
          For lCol = fg.FixedCols + 1 To fg.Cols
             If IsNumeric(fg.TextMatrix(lRow - 1, lCol - 1)) Then
                  excelDouble = CDbl(fg.TextMatrix(lRow - 1, lCol - 1)) + 0
                  .EWriteDouble lRow + rowOffset, lCol, excelDouble
             Else
                 .EWriteString lRow + rowOffset, lCol, fg.TextMatrix(lRow - 1, lCol - 1)
             End If
          Next lCol
        Next lRow
       
        
        .CloseFile
    End With
  
End Function

Private Sub AbortEdit()
    txtEdit.Visible = False
    fg.SetFocus
End Sub

Private Sub AutoNumerate()
 Dim i As Integer
  If m_AutoNumCol = True Then
   For i = 1 To fg.Rows - 1
    fg.TextMatrix(i, 0) = Str(i)
    fg.ColAlignment(0) = 3
   Next i
  End If
End Sub

Private Function AutoSizeC(myGrid As MSFlexGrid, _
                                Optional ByVal lfirstCol As Long = -1, _
                                Optional ByVal lLastCol As Long = -1, _
                                Optional bCheckFont As Boolean = False)
  
  Dim lCurCol As Long, lCurRow As Long
  Dim lCellWidth As Long, lColWidth As Long
  Dim bFontBold As Boolean
  Dim dFontSize As Double
  Dim sFontName As String
    
  bPrivateCellChange = True
  If bCheckFont Then
    ' save the forms font settings
    bFontBold = Me.FontBold
    sFontName = Me.FontName
    dFontSize = Me.FontSize
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
      .ColWidth(lCol) = lColWidth + UserControl.TextWidth("N")
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

Private Sub DisplayFormatedText(S As String, Row As Long, Col As Long)
    fg.TextMatrix(Row, Col) = Format(S, m_Cols(Col).ColDisplayFormat)
End Sub

Private Sub EndEdit(Cancel As Boolean)
    Dim nRow As Long
    Dim nCol As Long
    Dim sData
    sData = Split(txtEdit.Tag, "|")
    nRow = Val(sData(0))
    nCol = Val(sData(1))
    'show temporary
    Dim mOldText As String
    mOldText = fg.TextMatrix(nRow, nCol)
    bPrivateCellChange = True
    DisplayFormatedText txtEdit.Text, nRow, nCol
    bPrivateCellChange = False
    RaiseEvent ValidateEdit(nRow, nCol, Cancel)
    If Not Cancel Then
        'DisplayFormatedText txtEdit.Text, fg.Row, fg.Col
        DisplayFormatedText txtEdit.Text, nRow, nCol
        RaiseEvent AfterEdit(nRow, nCol)
        txtEdit.Visible = False
        fg.SetFocus
    Else
       ' If not validated then restore original text
       fg.TextMatrix(nRow, nCol) = mOldText
    End If
End Sub

Private Sub IsCellVisible()
    Dim a As Boolean
    a = fg.CellTop
End Sub

Private Sub LoadChkInCol(lCol As Long)
   Dim i As Integer
   
     With fg
         .Col = lCol
         For i = 1 To .Rows - 1
            .Row = i
            .CellPictureAlignment = 3
            Set .CellPicture = pUncheck.Picture
            '.TextMatrix(i, lCol) = i
         Next i
         .ColWidth(lCol) = 300
         .Refresh
     End With
     
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

Private Sub MoveCellOnEnter()
    If Not m_EnterKeyBehaviour = axEKNone Then
        If m_EnterKeyBehaviour = axEKMoveRight Then
            If fg.Col < fg.Cols - 1 Then
               fg.Col = fg.Col + 1
            Else
               If fg.Row < fg.Rows - 1 Then
                  Dim nRow As Long
                  Dim nCol As Long
                  nCol = fg.Col
                  nRow = fg.Row
                  nRow = fg.Row + 1
                  Dim i As Long
                  For i = fg.FixedCols To fg.Cols - 1
                      If fg.ColWidth(i) > 0 Then
                         nCol = i
                         Exit For
                      End If
                  Next
                  bPrivateCellChange = True
                  fg.Row = nRow
                  bPrivateCellChange = False
                  If fg.Col <> nCol Then
                     fg.Col = nCol
                  End If
               End If
            End If
            
        ElseIf m_EnterKeyBehaviour = axEKMoveDown Then
            If fg.Row < fg.Rows - 1 Then fg.Row = fg.Row + 1
        End If
        Call IsCellVisible
    End If
End Sub

Private Sub SetAlternateRowColors(lColor1 As Long, lColor2 As Long)
    Dim lOrgRow As Long, lOrgCol As Long
    Dim lColor As Long

    bPrivateCellChange = True

    With fg
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

Private Sub SetColumnObject(lCol As Long)
'Dim csText As String
'csText = fg.TextMatrix(oRow, oCol)

cBox.Visible = False
LBox.Visible = False
TBox.Visible = False
Boton.Visible = False

  Select Case m_Cols(lCol).ColType
    Case Is = eRemoveObject
        cBox.Visible = False
        LBox.Visible = False
        TBox.Visible = False
        Boton.Visible = False

    Case Is = eComboBoxColumn
        With fg
            If .Col < lCol Then Exit Sub
            cBox.Move .Left + .CellLeft - 2, .Top + .CellTop - 2, .CellWidth
        End With
        cBox.Text = fg.Text
        'cBox.ListIndex = 0
        cBox.Visible = True
        cBox.SetFocus
        
    Case Is = eListBoxColumn
        With fg
            If .Col < lCol Then Exit Sub
            LBox.Move .Left + .CellLeft - 2, .Top + .CellTop - 2, .CellWidth + 100, 1600
        End With
        'LBox.ListIndex = 0
        LBox.Text = fg.Text
        LBox.Visible = True
        LBox.SetFocus
        
    Case Is = eTextBoxColumn
        With fg
            If .Col < lCol Then Exit Sub
            TBox.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
            TBox.Text = .Text
        End With
        TBox.Visible = True
        TBox.SelStart = 0
        TBox.SelLength = Len(TBox)
        TBox.SetFocus
        Exit Sub

    Case Is = eButtonColumn
        With fg
            If .Col < lCol Then Exit Sub
            Boton.Move .Left + .CellLeft + .CellWidth, .Top + .CellTop, 230, .CellHeight
        End With
        Boton.Visible = True
        Exit Sub
         
  End Select
End Sub

Private Sub StartKeyEdit(KeyAscii As Integer, Optional bShowOldText As Boolean)
    If Not m_Editable Then Exit Sub
    With fg
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft - 2, .Top + .CellTop - 2, .CellWidth - 8, .CellHeight - 8
    End With
    
    Dim Cancel As Boolean
    RaiseEvent BeforeEdit(Cancel)
    
    If Not Cancel Then
        If bShowOldText Then
           txtEdit.Text = fg.Text
           txtEdit.SelStart = 0
           txtEdit.SelLength = Len(txtEdit.Text)
        Else
            txtEdit.Text = Chr$(KeyAscii)
            txtEdit.SelStart = 1
        End If
        txtEdit.Tag = fg.Row & "|" & fg.Col
        txtEdit.Visible = True
        txtEdit.SetFocus
    End If
End Sub

Private Sub Boton_Click()
RaiseEvent ButtonClick(fg.Row, fg.Col)
End Sub

Private Sub cBox_Click()
'fg.Text = cBox.Text
'cBox.Visible = False
bCboxSel = True
RaiseEvent ListClick(oRow, oCol, cBox.ListIndex)
End Sub

Private Sub cBox_DblClick()
On Error Resume Next
With fg
  .Text = cBox.Text
  cBox.Visible = False
  RaiseEvent ListClick(oRow, oCol, cBox.ListIndex)
  RaiseEvent CellTextChange(oRow, oCol)
End With
'Call MoveCellOnEnter
End Sub

Private Sub cBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    cBox.Visible = False
End If
End Sub

Private Sub fg_Click()
'On Error GoTo ErrSub0
Dim g As Integer

oRow = fg.Row
oCol = fg.Col
sTextOnCell = fg.TextMatrix(oRow, oCol)

With txtInfoBar
  .Text = "[R" & oRow & ":C" & oCol & "]:"
  Dim sRowText As String
  sRowText = ""
  
  Select Case m_SetInfoBar
    Case Is = CellGridInfo
      .Text = .Text & "<str>" & fg.TextMatrix(oRow, oCol) & " <val>" & CleanValue(fg.TextMatrix(oRow, oCol))
  
    Case Is = RowGridInfo
      ' Recorrer Toda la Fila del Grid
      For g = 0 To fg.Cols - 1
        sRowText = sRowText & "[" & fg.TextMatrix(oRow, g) & "]"
      Next g
      .Text = .Text & Mid$(sRowText, 1, Len(sRowText) - 1)
      
    Case Is = ColGridInfo
      ' Recorrer Toda la Columna del Grid
      For g = 0 To fg.Rows - 1
        sRowText = sRowText & "[" & fg.TextMatrix(g, oCol) & "]"
      Next g
      .Text = .Text & Mid$(sRowText, 1, Len(sRowText) - 1)
  End Select
    
End With

With fg
   If m_Cols(oCol).ColType = eCheckBoxColumn Then
      If .CellPicture = pUncheck Then
         Set .CellPicture = pCheck
      Else
         Set .CellPicture = pUncheck
      End If
   End If
End With

ErrSub0:
    RaiseEvent Click(oRow, oCol)
End Sub

Private Sub fg_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    RaiseEvent Compare(Row1, Row2, Cmp)
End Sub

Private Sub fg_DblClick()
Static bTipo As Boolean
If (fg.MouseRow = 0) Then
  ' Ordena en forma ascendente
  If bTipo Then
    fg.Col = fg.MouseCol
    fg.Sort = 2
    bTipo = False
    ' Ordena en forma descendente
  Else
    fg.Col = fg.MouseCol
    fg.Sort = 1
    bTipo = True
  End If
End If

    RaiseEvent DblClick(fg.Row, fg.Col)
End Sub

Private Sub fg_EnterCell()
    If Not bPrivateCellChange Then
       RaiseEvent EnterCell
    End If
    Call SetColumnObject(fg.Col)
End Sub

Private Sub fg_GotFocus()
    If txtEdit.Visible = True Then
        Dim Cancel As Boolean
        Call EndEdit(Cancel)
        If Cancel Then
           txtEdit.SetFocus
        End If
    End If
    HookScroll fg
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
If m_Editable Then
  If KeyCode = vbKeyF2 Then
      Call StartKeyEdit(0, True)
  ElseIf KeyCode = vbKeyDelete Then
    If fg.Row >= fg.FixedRows Then
      If MsgBox("Borrar Fila " & fg.Row & " ?" & vbNewLine & fg.TextMatrix(fg.Row, 1) & ", " & fg.TextMatrix(fg.Row, 1) & "...", vbYesNo + vbQuestion) = vbYes Then
         fg.RemoveItem (fg.Row) 'Delete the row
         KeyCode = 0 'prevent beep
         AutoNumerate 'actualiza numeración
      End If
    End If
  ElseIf KeyCode = vbKeyInsert Then
      If MsgBox("Insertar Fila en " & fg.Row & " ?", vbYesNo + vbQuestion) = vbYes Then
         Dim R As Integer, C As Integer
         Dim TotalRows As Integer, xRow As Integer
         With fg
             xRow = .Row
             TotalRows = .Rows - 1
             Debug.Print .Row
             .Rows = .Rows + 1
             ' Mover datos 1 fila hacia abajo
               For R = .Rows - 1 To xRow Step -1
                  For C = .Cols - 1 To 1 Step -1
                     .TextMatrix(R, C) = .TextMatrix(R - 1, C)
                  Next
               Next
           ' Limpiar Fila insertada
           If xRow = TotalRows Then
             For C = 1 To .Cols - 1
               .TextMatrix(xRow, C) = ""
             Next C
           Else
             For C = 1 To .Cols - 1
               .TextMatrix(xRow, C) = ""
             Next C
           End If
         End With
      End If
  End If
End If
End Sub


Private Sub fg_KeyPress(KeyAscii As Integer)
    Dim sInputMask As String
    sInputMask = m_Cols(fg.Col).ColInputMask
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
    RaiseEvent KeyPressEdit(fg.Row, fg.Col, KeyAscii)
    If KeyAscii = vbKeyTab Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        Call MoveCellOnEnter
        KeyAscii = 0
    End If
    If KeyAscii > 0 Then
        Call StartKeyEdit(KeyAscii)
    End If
End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub fg_LeaveCell()
    If Not bPrivateCellChange Then
       RaiseEvent LeaveCell
    End If
End Sub

Private Sub fg_LostFocus()
UnHookScroll fg
End Sub

Private Sub fG_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  oRow = fg.Row
  oCol = fg.Col
  Call SetColumnObject(fg.Col)
  '-----------------------
  Dim mRow As Integer, mCol As Integer
    mRow = fg.MouseRow
    mCol = fg.MouseCol
    If m_ShowMenuRow = True Then
      If Button = 2 And (mCol = 0 And mRow <> 0) Then
        fg.Col = IIf(mCol = 0, 1, mCol)
        fg.Row = IIf(mRow = 0, 1, mRow)
        PopupMenu MnuFGridRows
      End If
    End If
    If m_ShowMenuCol = True Then
      If Button = 2 And (mRow = 0 And mCol <> 0) Then
        fg.Col = IIf(mCol = 0, 1, mCol)
        fg.Row = IIf(mRow = 0, 1, mRow)
        PopupMenu MnuFGridCols
      End If
    End If
  '-----------------------
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub fg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub fg_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Static CurrentWidth As Single

  ' Check to see if the Cell's width has changed.
  If fg.CellWidth <> CurrentWidth Then
    cBox.Width = fg.CellWidth
    LBox.Width = fg.CellWidth + 270
    CurrentWidth = fg.CellWidth
  End If

   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub fg_OLECompleteDrag(Effect As Long)
   RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub fg_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub fg_OLEDragOver(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub fg_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub fg_OLESetData(Data As MSFlexGridLib.DataObject, DataFormat As Integer)
   RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub fg_OLEStartDrag(Data As MSFlexGridLib.DataObject, AllowedEffects As Long)
   RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub fg_RowColChange()
   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub fg_Scroll()
    If txtEdit.Visible = True Then
       AbortEdit
    End If
    
  cBox.Visible = False
  LBox.Visible = False
  TBox.Visible = False
  
    RaiseEvent Scroll
End Sub

Private Sub fg_SelChange()
   If Not bPrivateCellChange Then
      RaiseEvent SelChange
   End If
End Sub

Private Sub LBox_Click()
fg.Text = LBox.Text
LBox.Visible = False
RaiseEvent ListClick(fg.Row, fg.Col, LBox.ListIndex)
RaiseEvent CellTextChange(fg.Row, fg.Col)
'Call MoveCellOnEnter
End Sub

Private Sub LBox_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    LBox.Visible = False
    fg.Text = sTextOnCell
End If
End Sub

Private Sub mnuDeleteGridCol_Click()
    Dim C As Integer, n As Integer, x As Integer
    
    With fg
        If .Cols > 2 Then        'make sure we don't del a Col
          C = .Col
          For n = 1 To .Rows - 1
             If .TextMatrix(n, C) > "" Then
               x = 1
               Exit For
             End If
          Next
          If x Then
            n = MsgBox("Se quitará columna " & .Col & ": " & .TextMatrix(0, .Col) & vbNewLine & _
                                                             .TextMatrix(1, .Col) & vbNewLine & _
                                                             .TextMatrix(2, .Col) & vbNewLine & "...", vbYesNo, "Eliminar Columna " & Str$(.Col))
          End If
          If x = 0 Or n = 6 Then           'no exist. data or YES
            For n = .Col To .Cols - 2      'move exist data left 1 col
               For x = 0 To .Rows - 1
                  .TextMatrix(x, n) = .TextMatrix(x, n + 1)
               Next
            Next
            If C = .Cols - 1 Then     'set new cursor col
              .Col = .Cols - 2
            End If
            'Borra ultima Columna del Grid...
            .Cols = .Cols - 1
          End If
        End If
    End With

   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub mnuDeleteGridRow_Click()
    Dim R As Integer, n As Integer, x As Integer
    
    With fg
        If .Rows > 2 Then        'make sure we don't del a row
          R = .Row
          For n = 1 To .Cols - 1
             If .TextMatrix(R, n) > "" Then
               x = 1
               Exit For
             End If
          Next
          If x Then
            n = MsgBox("Se quitará linea " & .Row & ": " & .TextMatrix(R, 1) & "|" & _
                                                           .TextMatrix(R, 2) & "|" & _
                                                           .TextMatrix(R, 3) & "|...", vbYesNo, "Eliminar Linea " & Str$(R))
           End If
          If x = 0 Or n = 6 Then           'no exist. data or YES
            For n = .Row To .Rows - 2      'move exist data up 1 row
               For x = 1 To .Cols - 1
                  .TextMatrix(n, x) = .TextMatrix(n + 1, x)
               Next
            Next
            If R = .Rows - 1 Then     'set new cursor row
              .Row = .Rows - 2
            End If
            'Borra ultima Fila del Grid...
            .Rows = .Rows - 1
          End If
        End If
    End With
    
AutoNumerate

   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub MnuFGridAddCol_Click()
    Dim R As Integer, C As Integer
    Dim xTotalCols As Integer, xCol As Integer
    
    With fg
        xCol = .MouseCol
        xTotalCols = .Cols - 1
        Debug.Print .MouseCol
        .Cols = .Cols + 1
        ' Mover datos 1 columna a la derecha
          For C = .Cols - 1 To xCol Step -1
             For R = .Rows - 1 To 0 Step -1
                .TextMatrix(R, C) = .TextMatrix(R, C - 1)
             Next
          Next
      ' Limpiar columna insertada
      If xCol = xTotalCols Then
        For R = 0 To .Rows - 1
          .TextMatrix(R, xCol) = ""
        Next R
      Else
        For R = 0 To .Rows - 1
          .TextMatrix(R, xCol - 1) = ""
        Next R
      End If
    End With

   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub MnuFGridAddColEnd_Click()
fg.Cols = fg.Cols + 1

If Not bPrivateCellChange Then
    RaiseEvent RowColChange
End If
End Sub

Private Sub MnuFGridAddRow_Click()
    Dim R As Integer, C As Integer
    Dim xTotalRows As Integer, xRow As Integer
    
    With fg
        xRow = .MouseRow
        xTotalRows = .Rows - 1
        Debug.Print .MouseRow
        .Rows = .Rows + 1
        ' Mover datos 1 fila hacia abajo
          For R = .Rows - 1 To xRow Step -1
             For C = .Cols - 1 To 1 Step -1
                .TextMatrix(R, C) = .TextMatrix(R - 1, C)
             Next
          Next
      ' Limpiar Fila insertada
      If xRow = xTotalRows Then
        For C = 1 To .Cols - 1
          .TextMatrix(xRow, C) = ""
        Next C
      Else
        For C = 1 To .Cols - 1
          .TextMatrix(xRow - 1, C) = ""
        Next C
      End If
    End With

  AutoNumerate

   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub MnuFGridAddRowEnd_Click()
  fg.Rows = fg.Rows + 1
  AutoNumerate

   If Not bPrivateCellChange Then
       RaiseEvent RowColChange
   End If
End Sub

Private Sub TBox_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   fg.TextMatrix(fg.Row, fg.Col) = TBox.Text
   TBox.Visible = False
   KeyAscii = 0
   Call MoveCellOnEnter
   RaiseEvent CellTextChange(oRow, oCol)
ElseIf KeyAscii = vbKeyEscape Then
   TBox.Text = ""
   TBox.Visible = False
End If

End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sInputMask As String
sInputMask = m_Cols(fg.Col).ColInputMask

If LenB(sInputMask) > 0 Then
   NumKeyDown KeyCode, txtEdit, sInputMask
End If

RaiseEvent KeyDownEdit(fg.Row, fg.Col, KeyCode, Shift)
Dim Cancel As Boolean

If KeyCode = vbKeyDown Then
    Call EndEdit(Cancel)
    If Not Cancel Then
       If fg.Row < fg.Rows - 1 Then fg.Row = fg.Row + 1
    End If
End If

If KeyCode = vbKeyUp Then
    Call EndEdit(Cancel)
    If Not Cancel Then
        If fg.Row > fg.FixedRows Then fg.Row = fg.Row - 1
    End If
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim sInputMask As String
    sInputMask = m_Cols(fg.Col).ColInputMask
    If LenB(sInputMask) > 0 Then
       If Not KeyAscii = 13 Then
           NumKeyPress KeyAscii, txtEdit, sInputMask
           If KeyAscii = 0 Then
              Exit Sub
           End If
       End If
    End If
    RaiseEvent KeyPressEdit(fg.Row, fg.Col, KeyAscii)
    If KeyAscii = 13 Then
        Dim Cancel As Boolean
        Call EndEdit(Cancel)
        If Not Cancel Then Call MoveCellOnEnter
    End If
    If KeyAscii = 27 Then
        AbortEdit
    End If
End Sub

Private Sub txtEdit_LostFocus()
  Dim Cancel As Boolean
  Call EndEdit(Cancel)
End Sub

Private Sub UserControl_Initialize()
   ReDim m_Cols(fg.Cols)
End Sub

Private Sub UserControl_InitProperties()
m_SetInfoBar = RowGridInfo
End Sub

'*********************************
'Load property values from storage
'*********************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' New Properties
    m_AutoNumCol = PropBag.ReadProperty("AutoNumFirstCol", False)
    m_ShowMenuRow = PropBag.ReadProperty("ShowMenuRow", False)
    m_ShowMenuCol = PropBag.ReadProperty("ShowMenuCol", False)
    m_SetInfoBar = PropBag.ReadProperty("SetInfoBar", RowGridInfo)
    txtInfoBar.Visible = PropBag.ReadProperty("ShowInfoBar", False)
    m_EnterKeyBehaviour = PropBag.ReadProperty("EnterKeyBehaviour", axEKMoveRight)
    m_Editable = PropBag.ReadProperty("Editable", False)
    m_AutoSizeMode = PropBag.ReadProperty("AutoSizeMode", axAutoSizeColWidth)
    m_SortMode = PropBag.ReadProperty("SortColumnMode", flexSortNone)
    m_BackColorAlternate = PropBag.ReadProperty("BackColorAlternate", &H80000005)
    ' Mapped Properties
    fg.GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", flexGridInset)
    fg.GridLines = PropBag.ReadProperty("GridLines", flexGridFlat)
    fg.AllowBigSelection = PropBag.ReadProperty("AllowBigSelection", True)
    fg.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 0)
    fg.Appearance = PropBag.ReadProperty("Appearance", 1)
    fg.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    'Set Edit's backcolor similar to grid's backcolor
    txtEdit.BackColor = fg.BackColor
    fg.BackColorBkg = PropBag.ReadProperty("BackColorBkg", &H808080)
    fg.BackColorFixed = PropBag.ReadProperty("BackColorFixed", &H8000000F)
    fg.BackColorSel = PropBag.ReadProperty("BackColorSel", &H8000000D)
    fg.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    'fg.Cols = PropBag.ReadProperty("Cols", 2)
    Cols = PropBag.ReadProperty("Cols", 2)
    fg.Enabled = PropBag.ReadProperty("Enabled", True)
    fg.FillStyle = PropBag.ReadProperty("FillStyle", 0)
    fg.FixedCols = PropBag.ReadProperty("FixedCols", 1)
    fg.FixedRows = PropBag.ReadProperty("FixedRows", 1)
    fg.FocusRect = PropBag.ReadProperty("FocusRect", 1)
    Set fg.Font = PropBag.ReadProperty("Font", Ambient.Font)
    fg.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    fg.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", &H80000012)
    fg.ForeColorSel = PropBag.ReadProperty("ForeColorSel", &H8000000E)
    fg.FormatString = PropBag.ReadProperty("FormatString", "")
    fg.GridColor = PropBag.ReadProperty("GridColor", &HC0C0C0)
    fg.GridColorFixed = PropBag.ReadProperty("GridColorFixed", &H0&)
    fg.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
    fg.HighLight = PropBag.ReadProperty("HighLight", 1)
    fg.MergeCells = PropBag.ReadProperty("MergeCells", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    fg.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    fg.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    fg.PictureType = PropBag.ReadProperty("PictureType", 0)
    fg.Redraw = PropBag.ReadProperty("Redraw", True)
    fg.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    fg.RowHeightMin = PropBag.ReadProperty("RowHeightMin", cBox.Height)
    fg.Rows = PropBag.ReadProperty("Rows", 2)
    fg.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)
    fg.ScrollTrack = PropBag.ReadProperty("ScrollTrack", True)
    fg.SelectionMode = PropBag.ReadProperty("SelectionMode", 0)
    fg.Sort = PropBag.ReadProperty("Sort", 0)
    fg.TextStyle = PropBag.ReadProperty("TextStyle", 0)
    fg.TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", 0)
    fg.WordWrap = PropBag.ReadProperty("WordWrap", False)
    
    Set txtEdit.Font = fg.Font
    
    If m_BackColorAlternate <> fg.BackColor Then
       Call SetAlternateRowColors(m_BackColorAlternate, fg.BackColor)
    End If
    
    sDecimal = fGetLocaleInfo(LOCALE_SDECIMAL)
    sThousand = fGetLocaleInfo(LOCALE_SMONTHOUSANDSEP)
    sDateDiv = fGetLocaleInfo(LOCALE_SDATE)
    sMoney = fGetLocaleInfo(LOCALE_SCURRENCY)
    
End Sub

Private Sub UserControl_Resize()
With fg
    .Left = 0
    .Top = 0
    .Width = UserControl.Width
    ' Setting Rows MinHeight
    .RowHeightMin = cBox.Height
    If txtInfoBar.Visible = True Then
      .Height = UserControl.Height - 270
    Else
      .Height = UserControl.Height
    End If

End With
    
With txtInfoBar
  .Top = UserControl.ScaleHeight - .Height + 1
  .Left = 0
  .Width = UserControl.ScaleWidth
End With

End Sub

Private Sub UserControl_Terminate()
  If IsHooked Then
    Unhook   ' Stop checking messages.
  End If
  
  UnHookScroll fg
  
End Sub

'********************************
'Write property values to storage
'********************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' New Properties
    Call PropBag.WriteProperty("AutoNumFirstCol", m_AutoNumCol, False)
    Call PropBag.WriteProperty("ShowMenuRow", m_ShowMenuRow, False)
    Call PropBag.WriteProperty("ShowMenuCol", m_ShowMenuCol, False)
    Call PropBag.WriteProperty("SetInfoBar", m_SetInfoBar, RowGridInfo)
    Call PropBag.WriteProperty("ShowInfoBar", txtInfoBar.Visible, False)
    Call PropBag.WriteProperty("EnterKeyBehaviour", m_EnterKeyBehaviour, axEKMoveRight)
    Call PropBag.WriteProperty("Editable", m_Editable, False)
    Call PropBag.WriteProperty("AutoSizeMode", m_AutoSizeMode, axAutoSizeColWidth)
    Call PropBag.WriteProperty("SortColumnMode", m_SortMode, flexSortNone)
    Call PropBag.WriteProperty("BackColorAlternate", m_BackColorAlternate, &H80000005)
    ' Mapped Properties
    Call PropBag.WriteProperty("GridLines", fg.GridLines, 1)
    Call PropBag.WriteProperty("GridLinesFixed", fg.GridLinesFixed, 1)
    Call PropBag.WriteProperty("AllowBigSelection", fg.AllowBigSelection, True)
    Call PropBag.WriteProperty("AllowUserResizing", fg.AllowUserResizing, 0)
    Call PropBag.WriteProperty("Appearance", fg.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", fg.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColorBkg", fg.BackColorBkg, &H808080)
    Call PropBag.WriteProperty("BackColorFixed", fg.BackColorFixed, &H8000000F)
    Call PropBag.WriteProperty("BackColorSel", fg.BackColorSel, &H8000000D)
    Call PropBag.WriteProperty("BorderStyle", fg.BorderStyle, 1)
    Call PropBag.WriteProperty("Cols", fg.Cols, 2)
    Call PropBag.WriteProperty("Enabled", fg.Enabled, True)
    Call PropBag.WriteProperty("FillStyle", fg.FillStyle, 0)
    Call PropBag.WriteProperty("FixedCols", fg.FixedCols, 1)
    Call PropBag.WriteProperty("FixedRows", fg.FixedRows, 1)
    Call PropBag.WriteProperty("FocusRect", fg.FocusRect, 1)
    Call PropBag.WriteProperty("Font", fg.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", fg.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ForeColorFixed", fg.ForeColorFixed, &H80000012)
    Call PropBag.WriteProperty("ForeColorSel", fg.ForeColorSel, &H8000000E)
    Call PropBag.WriteProperty("FormatString", fg.FormatString, "")
    Call PropBag.WriteProperty("GridColor", fg.GridColor, &HC0C0C0)
    Call PropBag.WriteProperty("GridColorFixed", fg.GridColorFixed, &H0&)
    Call PropBag.WriteProperty("GridLineWidth", fg.GridLineWidth, 1)
    Call PropBag.WriteProperty("HighLight", fg.HighLight, 1)
    Call PropBag.WriteProperty("MergeCells", fg.MergeCells, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", fg.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", fg.OLEDropMode, 0)
    Call PropBag.WriteProperty("PictureType", fg.PictureType, 0)
    Call PropBag.WriteProperty("Redraw", fg.Redraw, True)
    Call PropBag.WriteProperty("RightToLeft", fg.RightToLeft, False)
    Call PropBag.WriteProperty("RowHeightMin", fg.RowHeightMin, cBox.Height)
    Call PropBag.WriteProperty("Rows", fg.Rows, 2)
    Call PropBag.WriteProperty("ScrollBars", fg.ScrollBars, 3)
    Call PropBag.WriteProperty("ScrollTrack", fg.ScrollTrack, True)
    Call PropBag.WriteProperty("SelectionMode", fg.SelectionMode, 0)
    Call PropBag.WriteProperty("TextStyle", fg.TextStyle, 0)
    Call PropBag.WriteProperty("TextStyleFixed", fg.TextStyleFixed, 0)
    Call PropBag.WriteProperty("WordWrap", fg.WordWrap, False)
        
End Sub


'**************************************
' Properties Mapped to Flexgrid Control
'**************************************

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,AllowBigSelection
Public Property Get AllowBigSelection() As Boolean
Attribute AllowBigSelection.VB_Description = "Returns/sets whether clicking on a column or row header should cause the entire column or row to be selected."
Attribute AllowBigSelection.VB_ProcData.VB_Invoke_Property = "General"
    AllowBigSelection = fg.AllowBigSelection
End Property

Public Property Let AllowBigSelection(ByVal New_AllowBigSelection As Boolean)
    fg.AllowBigSelection() = New_AllowBigSelection
    PropertyChanged "AllowBigSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,AllowUserResizing
Public Property Get AllowUserResizing() As AllowUserResizeSettings
Attribute AllowUserResizing.VB_Description = "Returns/sets whether the user should be allowed to resize rows and columns with the mouse."
Attribute AllowUserResizing.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowUserResizing = fg.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
    fg.AllowUserResizing() = New_AllowUserResizing
    PropertyChanged "AllowUserResizing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Appearance
Public Property Get Appearance() As AppearanceSettings
Attribute Appearance.VB_Description = "Returns/sets whether a control should be painted with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = fg.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceSettings)
    If New_Appearance = flex3D Then
    fg.Appearance() = 1
    cFlatSb.UninitializeCoolSB
    txtInfoBar.Appearance = 1
Else
    fg.Appearance() = 0
    cFlatSb.InitializeCoolSB fg.hwnd, False
    txtInfoBar.Appearance = 0
End If

    PropertyChanged "Appearance"
End Property

Public Property Get AutoNumFirstCol() As Boolean
  AutoNumFirstCol = m_AutoNumCol
End Property

Public Property Let AutoNumFirstCol(bAutoNum As Boolean)
  m_AutoNumCol = bAutoNum
  'llamar a procedimiento AutoNumerante
  Call AutoNumerate
  PropertyChanged "AutoNumFirstCol"
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
    If m_BackColorAlternate <> fg.BackColor Then
       Call SetAlternateRowColors(m_BackColorAlternate, fg.BackColor)
    End If
    PropertyChanged "BackColorAlternate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColorBkg
Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns/sets the background color of various elements of the FlexGrid."
Attribute BackColorBkg.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorBkg = fg.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    fg.BackColorBkg() = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColorFixed
Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "Returns/sets the background color of various elements of the FlexGrid."
Attribute BackColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorFixed = fg.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
    fg.BackColorFixed() = New_BackColorFixed
    PropertyChanged "BackColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of various elements of the FlexGrid."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = fg.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If txtEdit.BackColor = fg.BackColor Then
       txtEdit.BackColor = New_BackColor
    End If

    fg.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Dim OldBackColorAlternate As Long
    OldBackColorAlternate = m_BackColorAlternate
    BackColorAlternate = New_BackColor
    If m_BackColorAlternate <> OldBackColorAlternate Then
       Call SetAlternateRowColors(m_BackColorAlternate, m_BackColorAlternate)
    End If

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BackColorSel
Public Property Get BackColorSel() As OLE_COLOR
Attribute BackColorSel.VB_Description = "Returns/sets the background color of various elements of the FlexGrid."
Attribute BackColorSel.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorSel = fg.BackColorSel
End Property

Public Property Let BackColorSel(ByVal New_BackColorSel As OLE_COLOR)
    fg.BackColorSel() = New_BackColorSel
    PropertyChanged "BackColorSel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleSettings
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = fg.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    fg.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get CalculateColumn(Settings As eSubTotalSettings, eColumn As Long, _
                                    eRowInitial As Long, eRowFinal As Long) As Double
         bPrivateCellChange = True
         Dim nValue As Double
         nValue = 0
         Dim bFirst As Boolean
         bFirst = True
         Dim i As Long
         For i = eRowInitial To eRowFinal
            Select Case Settings
                Case Is = axSTSum
                     nValue = nValue + CleanValue(fg.TextMatrix(i, eColumn))
                Case Is = axSTMax
                     If bFirst Then
                        nValue = CleanValue(fg.TextMatrix(i, eColumn))
                        bFirst = False
                     Else
                        nValue = Max(nValue, CleanValue(fg.TextMatrix(i, eColumn)))
                     End If
                Case Is = axSTMin
                     If bFirst Then
                        nValue = CleanValue(fg.TextMatrix(i, eColumn))
                        bFirst = False
                     Else
                        nValue = Min(nValue, CleanValue(fg.TextMatrix(i, eColumn)))
                     End If
                Case Is = axSTCount
                        nValue = nValue + IIf(Len(fg.TextMatrix(i, eColumn)) > 0, 1, 0)
            End Select
         Next i
         
         If Settings = axSTMultiply Then
            nValue = CleanValue(fg.TextMatrix(eRowInitial, eColumn)) * CleanValue(fg.TextMatrix(eRowFinal, eColumn))
         End If
         
         bPrivateCellChange = False
        CalculateColumn = nValue
End Property

Public Property Get CalculateMatrix(Settings As eSubTotalSettings, _
                     eRowInitial As Long, eColumnInitial As Long, _
                       eRowFinal As Long, eColumnFinal As Long) As Double
         bPrivateCellChange = True
         Dim nValue As Double
         nValue = 0
         Dim bFirst As Boolean
         bFirst = True
         Dim i As Long
         Dim j As Long
         For i = eRowInitial To eRowFinal
             For j = eColumnInitial To eColumnFinal
                 Select Case Settings
                        Case Is = axSTSum
                              nValue = nValue + CleanValue(fg.TextMatrix(i, j))
                        Case Is = axSTMax
                             If bFirst Then
                                nValue = CleanValue(fg.TextMatrix(i, j))
                                bFirst = False
                             Else
                                nValue = Max(nValue, CleanValue(fg.TextMatrix(i, j)))
                             End If
                        Case Is = axSTMin
                             If bFirst Then
                                nValue = CleanValue(fg.TextMatrix(i, j))
                                bFirst = False
                             Else
                                nValue = Min(nValue, CleanValue(fg.TextMatrix(i, j)))
                             End If
                        Case Is = axSTCount
                             nValue = nValue + IIf(Len(fg.TextMatrix(i, j)) > 0, 1, 0)
                 End Select
             Next j
         Next i
         bPrivateCellChange = False
        CalculateMatrix = nValue
End Property

Public Property Get CalculateRow(Settings As eSubTotalSettings, eRow As Long, _
                                    eColInitial As Long, eColFinal As Long) As Double
         bPrivateCellChange = True
         Dim nValue As Double
         nValue = 0
         Dim bFirst As Boolean
         bFirst = True
         Dim i As Long
         For i = eColInitial To eColFinal
            Select Case Settings
                Case Is = axSTSum
                     nValue = nValue + CleanValue(fg.TextMatrix(eRow, i))
                Case Is = axSTMax
                     If bFirst Then
                        nValue = CleanValue(fg.TextMatrix(eRow, i))
                        bFirst = False
                     Else
                        nValue = Max(nValue, CleanValue(fg.TextMatrix(eRow, i)))
                     End If
                Case Is = axSTMin
                     If bFirst Then
                        nValue = CleanValue(fg.TextMatrix(eRow, i))
                        bFirst = False
                     Else
                        nValue = Min(nValue, CleanValue(fg.TextMatrix(eRow, i)))
                     End If
                Case Is = axSTCount
                        nValue = nValue + IIf(Len(fg.TextMatrix(eRow, i)) > 0, 1, 0)
            End Select
         Next i
         
         If Settings = axSTMultiply Then
            nValue = Val(fg.TextMatrix(eRow, eColInitial)) * CleanValue(fg.TextMatrix(eRow, eColFinal))
         End If
         
         bPrivateCellChange = False
        CalculateRow = nValue
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellAlignment() As AlignmentSettings
Attribute CellAlignment.VB_MemberFlags = "400"
    CellAlignment = fg.CellAlignment
End Property

Public Property Let CellAlignment(ByVal New_CellAlignment As AlignmentSettings)
    fg.CellAlignment = New_CellAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellBackColor() As OLE_COLOR
Attribute CellBackColor.VB_MemberFlags = "400"
    CellBackColor = fg.CellBackColor
End Property

Public Property Let CellBackColor(ByVal New_CellBackColor As OLE_COLOR)
    fg.CellBackColor = New_CellBackColor
End Property

Public Property Get CellCheckValue(ByVal Row As Long, ByVal Col As Long) As Boolean
 With fg
    .Row = Row
    .Col = Col
   If m_Cols(Col).ColType = eCheckBoxColumn Then
      If .CellPicture = pCheck Then
         CellCheckValue = True
      Else
         CellCheckValue = False
      End If
   End If
 End With
   
End Property

Public Property Let CellCheckValue(ByVal Row As Long, ByVal Col As Long, chkValue As Boolean)
With fg
    .Row = Row
    .Col = Col
   If m_Cols(Col).ColType = eCheckBoxColumn Then
      If chkValue = True Then
         Set .CellPicture = pCheck
      Else
         Set .CellPicture = pUncheck
      End If
   End If
End With

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontBold() As Boolean
Attribute CellFontBold.VB_MemberFlags = "400"
    CellFontBold = fg.CellFontBold
End Property

Public Property Let CellFontBold(ByVal New_CellFontBold As Boolean)
    fg.CellFontBold = New_CellFontBold
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontItalic() As Boolean
Attribute CellFontItalic.VB_MemberFlags = "400"
    CellFontItalic = fg.CellFontItalic
End Property

Public Property Let CellFontItalic(ByVal New_CellFontItalic As Boolean)
    fg.CellFontItalic = New_CellFontItalic
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontName() As String
Attribute CellFontName.VB_MemberFlags = "400"
    CellFontName = fg.CellFontName
End Property

Public Property Let CellFontName(ByVal New_CellFontName As String)
    fg.CellFontName = New_CellFontName
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontSize() As Single
Attribute CellFontSize.VB_MemberFlags = "400"
    CellFontSize = fg.CellFontSize
End Property

Public Property Let CellFontSize(ByVal New_CellFontSize As Single)
    fg.CellFontSize = New_CellFontSize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontStrikeThrough() As Boolean
Attribute CellFontStrikeThrough.VB_MemberFlags = "400"
    CellFontStrikeThrough = fg.CellFontStrikeThrough
End Property

Public Property Let CellFontStrikeThrough(ByVal New_CellFontStrikeThrough As Boolean)
    fg.CellFontStrikeThrough = New_CellFontStrikeThrough
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontUnderline() As Boolean
Attribute CellFontUnderline.VB_MemberFlags = "400"
    CellFontUnderline = fg.CellFontUnderline
End Property

Public Property Let CellFontUnderline(ByVal New_CellFontUnderline As Boolean)
    fg.CellFontUnderline = New_CellFontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellFontWidth() As Single
Attribute CellFontWidth.VB_MemberFlags = "400"
    CellFontWidth = fg.CellFontWidth
End Property

Public Property Let CellFontWidth(ByVal New_CellFontWidth As Single)
    fg.CellFontWidth = New_CellFontWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellForeColor() As OLE_COLOR
Attribute CellForeColor.VB_MemberFlags = "400"
    CellForeColor = fg.CellForeColor
End Property

Public Property Let CellForeColor(ByVal New_CellForeColor As OLE_COLOR)
    fg.CellForeColor = New_CellForeColor
End Property

Public Property Get cell(Setting As eCellProperty, _
            ByVal Row1 As Long, ByVal Col1 As Long, _
            ByVal Row2 As Long, ByVal Col2 As Long) As Variant
    bPrivateCellChange = True
    Dim OldRow As Long
    Dim OldCol As Long
    OldRow = fg.Row
    OldCol = fg.Col
    fg.Row = Row1
    fg.Col = Col1
    Select Case Setting
        Case Is = axcpCellAlignment
            cell = fg.CellAlignment
        Case Is = axcpCellFontBold
            cell = fg.CellFontBold
        Case Is = axcpCellFontName
            cell = fg.CellFontName
        Case Is = axcpCellFontSize
            cell = fg.CellFontSize
        Case Is = axcpCellForeColor
            cell = fg.CellForeColor
        Case Is = axcpCellBackColor
            cell = fg.CellBackColor
    End Select
    fg.Row = OldRow
    fg.Col = OldCol
    bPrivateCellChange = False
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellHeight
Public Property Get CellHeight() As Long
    CellHeight = fg.CellHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellLeft
Public Property Get CellLeft() As Long
    CellLeft = fg.CellLeft
End Property

Public Property Let cell(Setting As eCellProperty, _
            ByVal Row1 As Long, ByVal Col1 As Long, _
            ByVal Row2 As Long, ByVal Col2 As Long, _
            New_Val As Variant)
    bPrivateCellChange = True
    Dim OldRow As Long
    Dim OldCol As Long
    OldRow = fg.Row
    OldCol = fg.Col
    Dim i As Long
    Dim j As Long
    For i = Row1 To Row2
        For j = Col1 To Col2
            fg.Row = i
            fg.Col = j
            Select Case Setting
                Case Is = axcpCellAlignment
                    fg.CellAlignment = New_Val
                Case Is = axcpCellFontBold
                    fg.CellFontBold = New_Val
                Case Is = axcpCellFontName
                    fg.CellFontName = New_Val
                Case Is = axcpCellFontSize
                    fg.CellFontSize = New_Val
                Case Is = axcpCellForeColor
                    fg.CellForeColor = New_Val
                Case Is = axcpCellBackColor
                    fg.CellBackColor = New_Val
            End Select
        Next j
    Next i
    fg.Row = OldRow
    fg.Col = OldCol
    bPrivateCellChange = False
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellPictureAlignment() As AlignmentSettings
Attribute CellPictureAlignment.VB_MemberFlags = "400"
    CellPictureAlignment = fg.CellPictureAlignment
End Property

Public Property Let CellPictureAlignment(ByVal New_CellPictureAlignment As AlignmentSettings)
    fg.CellPictureAlignment = New_CellPictureAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellPicture
Public Property Get CellPicture() As Picture
Attribute CellPicture.VB_Description = "Returns/sets an image to be displayed in the current cell or in a range of cells."
Attribute CellPicture.VB_MemberFlags = "400"
    Set CellPicture = fg.CellPicture
End Property

Public Property Set CellPicture(ByVal New_CellPicture As Picture)
    Set fg.CellPicture = New_CellPicture
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CellTextStyle() As TextStyleSettings
Attribute CellTextStyle.VB_MemberFlags = "400"
    CellTextStyle = fg.CellTextStyle
End Property

Public Property Let CellTextStyle(ByVal New_CellTextStyle As TextStyleSettings)
    fg.CellTextStyle = New_CellTextStyle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellTop
Public Property Get CellTop() As Long
    CellTop = fg.CellTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,CellWidth
Public Property Get CellWidth() As Long
    CellWidth = fg.CellWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Clip() As String
Attribute Clip.VB_MemberFlags = "400"
    Clip = fg.Clip
End Property

Public Property Let Clip(ByVal New_Clip As String)
    fg.Clip = New_Clip
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColAlignment
Public Property Get ColAlignment(ByVal Index As Long) As eAlignCols
Attribute ColAlignment.VB_Description = "Returns/sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
    ColAlignment = fg.ColAlignment(Index)
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal New_ColAlignment As eAlignCols)
    fg.ColAlignment(Index) = New_ColAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColData
Public Property Get ColData(ByVal Index As Long) As Long
Attribute ColData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the FlexGrid. Not available at design time."
    ColData = fg.ColData(Index)
End Property

Public Property Let ColData(ByVal Index As Long, ByVal New_ColData As Long)
    fg.ColData(Index) = New_ColData
End Property

Public Property Get ColDisplayFormat(ByVal Col As Long) As String
    ColDisplayFormat = m_Cols(Col).ColDisplayFormat
End Property

Public Property Let ColDisplayFormat(ByVal Col As Long, ByVal New_ColDisplayFormat As String)
    m_Cols(Col).ColDisplayFormat = New_ColDisplayFormat
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Col() As Long
Attribute Col.VB_MemberFlags = "400"
    Col = fg.Col
End Property

Public Property Get ColInputMask(ByVal Col As Long) As String
    ColInputMask = m_Cols(Col).ColInputMask
End Property

Public Property Let ColInputMask(ByVal Col As Long, ByVal New_ColInputMask As String)
    m_Cols(Col).ColInputMask = New_ColInputMask
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColIsVisible
Public Property Get ColIsVisible(ByVal Index As Long) As Boolean
    ColIsVisible = fg.ColIsVisible(Index)
End Property

Public Property Let Col(ByVal New_Col As Long)
    fg.Col = New_Col
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColPos
Public Property Get ColPos(ByVal Index As Long) As Long
    ColPos = fg.ColPos(Index)
End Property

Public Property Let ColPosition(ByVal Index As Long, ByVal New_ColPosition As Long)
Attribute ColPosition.VB_Description = "Returns the distance in Twips between the upper-left corner of the control and the upper-left corner of a specified column."
    fg.ColPosition(Index) = New_ColPosition
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ColSel() As Long
Attribute ColSel.VB_MemberFlags = "400"
    ColSel = fg.ColSel
End Property

Public Property Let ColSel(ByVal New_ColSel As Long)
    fg.ColSel = New_ColSel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Cols
Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Determines the total number of columns or rows in a FlexGrid."
Attribute Cols.VB_ProcData.VB_Invoke_Property = "General"
    Cols = fg.Cols
End Property

Public Property Let Cols(ByVal New_Cols As Long)
    On Error GoTo ErrHand
    fg.Cols() = New_Cols
    ReDim Preserve m_Cols(New_Cols + 1)

    PropertyChanged "Cols"
    Exit Property
ErrHand:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ColWidth
Public Property Get ColWidth(ByVal Index As Long) As Long
Attribute ColWidth.VB_Description = "Determines the width of the specified column in Twips. Not available at design time."
    ColWidth = fg.ColWidth(Index)
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal New_ColWidth As Long)
    fg.ColWidth(Index) = New_ColWidth
End Property

Public Property Get Editable() As Boolean
Attribute Editable.VB_Description = "Returns/sets a value that determines whether data in grid can be edit"
Attribute Editable.VB_ProcData.VB_Invoke_Property = "Additional"
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
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = fg.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    fg.Enabled() = New_Enabled
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
Public Property Get FillStyle() As FillStyleSettings
Attribute FillStyle.VB_Description = "Determines whether setting the Text property or one of the Cell formatting properties of a FlexGrid applies the change to all selected cells."
    FillStyle = fg.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleSettings)
    fg.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FixedAlignment
Public Property Get FixedAlignment(ByVal Index As Long) As Integer
Attribute FixedAlignment.VB_Description = "Returns/sets the alignment of data in the fixed cells of a column."
    FixedAlignment = fg.FixedAlignment(Index)
End Property

Public Property Let FixedAlignment(ByVal Index As Long, ByVal New_FixedAlignment As Integer)
    fg.FixedAlignment(Index) = New_FixedAlignment
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FixedCols
Public Property Get FixedCols() As Long
Attribute FixedCols.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for a FlexGrid."
Attribute FixedCols.VB_ProcData.VB_Invoke_Property = "General"
    FixedCols = fg.FixedCols
End Property

Public Property Let FixedCols(ByVal New_FixedCols As Long)
    fg.FixedCols() = New_FixedCols
    PropertyChanged "FixedCols"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FixedRows
Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "Returns/sets the total number of fixed (non-scrollable) columns or rows for a FlexGrid."
Attribute FixedRows.VB_ProcData.VB_Invoke_Property = "General"
    FixedRows = fg.FixedRows
End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)
    fg.FixedRows() = New_FixedRows
    PropertyChanged "FixedRows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FocusRect
Public Property Get FocusRect() As FocusRectSettings
Attribute FocusRect.VB_Description = "Determines whether the FlexGrid control should draw a focus rectangle around the current cell."
    FocusRect = fg.FocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As FocusRectSettings)
    fg.FocusRect() = New_FocusRect
    PropertyChanged "FocusRect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "440"
    FontBold = fg.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    fg.FontBold = New_FontBold
    txtEdit.FontBold = New_FontBold
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets the default font or the font for individual cells."
Attribute Font.VB_UserMemId = -512
    Set Font = fg.Font
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_MemberFlags = "40"
    FontItalic = fg.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    fg.FontItalic = New_FontItalic
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "40"
    FontName = fg.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    fg.FontName = New_FontName
    txtEdit.FontName = New_FontName
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set fg.Font = New_Font
    Set txtEdit.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "40"
    FontSize = fg.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Long)
    fg.FontSize = New_FontSize
    txtEdit.FontSize = New_FontSize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_MemberFlags = "40"
    FontStrikethru = fg.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    fg.FontStrikethru = New_FontStrikethru
    txtEdit.FontStrikethru = New_FontStrikethru
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "40"
    FontUnderline = fg.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    fg.FontUnderline = New_FontUnderline
    txtEdit.FontUnderline = New_FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FontWidth
Public Property Get FontWidth() As Single
Attribute FontWidth.VB_Description = "Returns or sets the width, in points, of the font to be used for text displayed."
Attribute FontWidth.VB_MemberFlags = "400"
    FontWidth = fg.FontWidth
End Property

Public Property Let FontWidth(ByVal New_FontWidth As Single)
    fg.FontWidth() = New_FontWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ForeColorFixed
Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "Determines the color used to draw text on each part of the FlexGrid."
Attribute ForeColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorFixed = fg.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    fg.ForeColorFixed() = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Determines the color used to draw text on each part of the FlexGrid."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = fg.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    fg.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ForeColorSel
Public Property Get ForeColorSel() As OLE_COLOR
Attribute ForeColorSel.VB_Description = "Determines the color used to draw text on each part of the FlexGrid."
Attribute ForeColorSel.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorSel = fg.ForeColorSel
End Property

Public Property Let ForeColorSel(ByVal New_ForeColorSel As OLE_COLOR)
    fg.ForeColorSel() = New_ForeColorSel
    PropertyChanged "ForeColorSel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,FormatString
Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "Allows you to set up a FlexGrid's column widths, alignments, and fixed row and column text at design time. See the help file for details."
Attribute FormatString.VB_ProcData.VB_Invoke_Property = "Style"
    FormatString = fg.FormatString
End Property

Public Property Let FormatString(ByVal New_FormatString As String)
    fg.FormatString() = New_FormatString
    PropertyChanged "FormatString"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,GridColorFixed
Public Property Get GridColorFixed() As OLE_COLOR
Attribute GridColorFixed.VB_Description = "Returns/sets the color used to draw the lines between FlexGrid cells."
Attribute GridColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridColorFixed = fg.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
    fg.GridColorFixed() = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,GridColor
Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns/sets the color used to draw the lines between FlexGrid cells."
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridColor = fg.GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    fg.GridColor() = New_GridColor
    PropertyChanged "GridColor"
End Property

Public Property Get GridLinesFixed() As GridLineSettings
    GridLinesFixed = fg.GridLineWidth
End Property

Public Property Let GridLinesFixed(ByVal New_GridLinesFixed As GridLineSettings)
    fg.GridLinesFixed() = New_GridLinesFixed
    PropertyChanged "GridLinesFixed"
End Property

Public Property Get GridLines() As GridLineSettings
   GridLines = fg.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As GridLineSettings)
  fg.GridLines = New_GridLines
  PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,GridLineWidth
Public Property Get GridLineWidth() As Integer
Attribute GridLineWidth.VB_Description = "Returns/sets the width in Pixels of the gridlines for the control."
    GridLineWidth = fg.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Integer)
    fg.GridLineWidth() = New_GridLineWidth
    PropertyChanged "GridLineWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,HighLight
Public Property Get HighLight() As HighLightSettings
Attribute HighLight.VB_Description = "Returns/sets whether selected cells appear highlighted."
    HighLight = fg.HighLight
End Property

Public Property Let HighLight(ByVal New_HighLight As HighLightSettings)
    fg.HighLight() = New_HighLight
    PropertyChanged "HighLight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get hwnd() As Long
    hwnd = fg.hwnd
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
Public Property Get LeftCol() As Long
Attribute LeftCol.VB_MemberFlags = "400"
    LeftCol = fg.LeftCol
End Property

Public Property Let LeftCol(ByVal New_LeftCol As Long)
    fg.LeftCol = New_LeftCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MergeCells
Public Property Get MergeCells() As MergeCellsSettings
Attribute MergeCells.VB_Description = "Returns/sets whether cells with the same contents should be grouped in a single cell spanning multiple rows or columns."
    MergeCells = fg.MergeCells
End Property

Public Property Let MergeCells(ByVal New_MergeCells As MergeCellsSettings)
    fg.MergeCells() = New_MergeCells
    PropertyChanged "MergeCells"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MergeCol
Public Property Get MergeCol(ByVal Index As Long) As Boolean
Attribute MergeCol.VB_Description = "Returns/sets which rows (columns) should have their contents merged when the MergeCells property is set to a value other than 0 - Never."
Attribute MergeCol.VB_MemberFlags = "400"
    MergeCol = fg.MergeCol(Index)
End Property

Public Property Let MergeCol(ByVal Index As Long, ByVal New_MergeCol As Boolean)
    fg.MergeCol(Index) = New_MergeCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MergeRow
Public Property Get MergeRow(ByVal Index As Long) As Boolean
Attribute MergeRow.VB_Description = "Returns/sets which rows (columns) should have their contents merged when the MergeCells property is set to a value other than 0 - Never."
Attribute MergeRow.VB_MemberFlags = "400"
    MergeRow = fg.MergeRow(Index)
End Property

Public Property Let MergeRow(ByVal Index As Long, ByVal New_MergeRow As Boolean)
    fg.MergeRow(Index) = New_MergeRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MouseCol
Public Property Get MouseCol() As Long
    MouseCol = fg.MouseCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
    Set MouseIcon = fg.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set fg.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MousePointer
Public Property Get MousePointer() As MousePointerSettings
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = fg.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerSettings)
    fg.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,MouseRow
Public Property Get MouseRow() As Long
    MouseRow = fg.MouseRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this control can act as an OLE drop target."
    OLEDropMode = fg.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    fg.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = fg.Picture
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,PictureType
Public Property Get PictureType() As PictureTypeSettings
Attribute PictureType.VB_Description = "Returns/sets the type of picture that should be generated by the Picture property."
    PictureType = fg.PictureType
End Property

Public Property Let PictureType(ByVal New_PictureType As PictureTypeSettings)
    fg.PictureType() = New_PictureType
    PropertyChanged "PictureType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Redraw
Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Enables or disables redrawing of the FlexGrid control."
    Redraw = fg.Redraw
End Property

Public Property Let Redraw(ByVal New_Redraw As Boolean)
    fg.Redraw() = New_Redraw
    PropertyChanged "Redraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = fg.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    fg.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowData
Public Property Get RowData(ByVal Index As Long) As Long
Attribute RowData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the FlexGrid. Not available at design time."
Attribute RowData.VB_MemberFlags = "400"
    RowData = fg.RowData(Index)
End Property

Public Property Let RowData(ByVal Index As Long, ByVal New_RowData As Long)
    fg.RowData(Index) = New_RowData
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Row() As Long
Attribute Row.VB_MemberFlags = "400"
    Row = fg.Row
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowHeight
Public Property Get RowHeight(ByVal Index As Long) As Long
Attribute RowHeight.VB_Description = "Returns/sets the height of the specified row in Twips. Not available at design time."
Attribute RowHeight.VB_MemberFlags = "400"
    RowHeight = fg.RowHeight(Index)
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal New_RowHeight As Long)
    fg.RowHeight(Index) = New_RowHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowHeightMin
Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Returns/sets a minimum row height for the entire control, in Twips."
Attribute RowHeightMin.VB_ProcData.VB_Invoke_Property = "Style"
    RowHeightMin = fg.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    fg.RowHeightMin() = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowIsVisible
Public Property Get RowIsVisible(ByVal Index As Long) As Boolean
    RowIsVisible = fg.RowIsVisible(Index)
End Property

Public Property Let Row(ByVal New_Row As Long)
    fg.Row = New_Row
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,RowPos
Public Property Get RowPos(ByVal Index As Long) As Long
    RowPos = fg.RowPos(Index)
End Property

Public Property Let RowPosition(ByVal Index As Long, ByVal New_RowPosition As Long)
Attribute RowPosition.VB_Description = "Returns the distance in Twips between the upper-left corner of the control and the upper-left corner of a specified row."
Attribute RowPosition.VB_MemberFlags = "400"
    fg.RowPosition(Index) = New_RowPosition
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get RowSel() As Long
Attribute RowSel.VB_MemberFlags = "400"
    RowSel = fg.RowSel
End Property

Public Property Let RowSel(ByVal New_RowSel As Long)
    fg.RowSel = New_RowSel
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,Rows
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in a FlexGrid."
Attribute Rows.VB_ProcData.VB_Invoke_Property = "General"
    Rows = fg.Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
    Dim mOldRows As Long
    mOldRows = fg.Rows
    fg.Rows() = New_Rows
        
    If New_Rows > mOldRows Then
        If m_BackColorAlternate <> fg.BackColor Then
            Call SetAlternateRowColors(m_BackColorAlternate, fg.BackColor)
        End If
    End If
    AutoNumerate
    PropertyChanged "Rows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ScrollBars
Public Property Get ScrollBars() As ScrollBarsSettings
Attribute ScrollBars.VB_Description = "Returns/sets whether a FlexGrid has horizontal or vertical scroll bars."
    ScrollBars = fg.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
    fg.ScrollBars() = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,ScrollTrack
Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns/sets whether FlexGrid should scroll its contents while the user moves the scroll box along the scroll bars."
    ScrollTrack = fg.ScrollTrack
End Property

Public Property Let ScrollTrack(ByVal New_ScrollTrack As Boolean)
    fg.ScrollTrack() = New_ScrollTrack
    PropertyChanged "ScrollTrack"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,SelectionMode
Public Property Get SelectionMode() As SelectionModeSettings
Attribute SelectionMode.VB_Description = "Returns/sets whether a FlexGrid should allow regular cell selection, selection by rows, or selection by columns."
    SelectionMode = fg.SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
    fg.SelectionMode() = New_SelectionMode
    PropertyChanged "SelectionMode"
End Property

Public Property Let SetColObject(ByVal Col As Long, ByVal new_ColType As eColumnType)
   If new_ColType = eCheckBoxColumn Then
      Call LoadChkInCol(Col)
   End If
   
   m_Cols(Col).ColType = new_ColType
End Property

Public Property Get SetInfoBar() As eTypeSingleInfoBar
    SetInfoBar = m_SetInfoBar
End Property

Public Property Let SetInfoBar(eNewType As eTypeSingleInfoBar)
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

Public Property Get ShowMenuCol() As Boolean
  ShowMenuCol = m_ShowMenuCol
End Property

Public Property Let ShowMenuCol(bShowMenuC As Boolean)
  m_ShowMenuCol = bShowMenuC
  PropertyChanged "ShowMenuCol"
End Property

Public Property Get ShowMenuRow() As Boolean
  ShowMenuRow = m_ShowMenuRow
End Property

Public Property Let ShowMenuRow(bShowMenuR As Boolean)
  m_ShowMenuRow = bShowMenuR
  PropertyChanged "ShowMenuRow"
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
Public Property Let Sort(ByVal New_Sort As SortSettings)
    fg.Sort() = New_Sort
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextArray
Public Property Get TextArray(ByVal Index As Long) As String
Attribute TextArray.VB_Description = "Returns/sets the text contents of an arbitrary cell (single subscript)."
Attribute TextArray.VB_MemberFlags = "400"
    TextArray = fg.TextArray(Index)
End Property

Public Property Let TextArray(ByVal Index As Long, ByVal New_TextArray As String)
    fg.TextArray(Index) = New_TextArray
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
    Text = fg.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Call DisplayFormatedText(New_Text, fg.Row, fg.Col)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextMatrix
Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns/sets the text contents of an arbitrary cell (row/col subscripts)."
Attribute TextMatrix.VB_MemberFlags = "400"
    TextMatrix = fg.TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
    Call DisplayFormatedText(New_TextMatrix, Row, Col)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextStyleFixed
Public Property Get TextStyleFixed() As TextStyleSettings
Attribute TextStyleFixed.VB_Description = "Returns/sets 3D effects for displaying text."
    TextStyleFixed = fg.TextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal New_TextStyleFixed As TextStyleSettings)
    fg.TextStyleFixed() = New_TextStyleFixed
    PropertyChanged "TextStyleFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,TextStyle
Public Property Get TextStyle() As TextStyleSettings
Attribute TextStyle.VB_Description = "Returns/sets 3D effects for displaying text."
    TextStyle = fg.TextStyle
End Property

Public Property Let TextStyle(ByVal New_TextStyle As TextStyleSettings)
    fg.TextStyle() = New_TextStyle
    PropertyChanged "TextStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get TopRow() As Long
Attribute TopRow.VB_MemberFlags = "400"
    TopRow = fg.TopRow
End Property

Public Property Let TopRow(ByVal New_TopRow As Long)
    fg.TopRow = New_TopRow
End Property

Public Property Get Value() As Variant
     Value = Val(fg.Text)
End Property

Public Property Get ValueMatrix(ByVal Row As Long, ByVal Col As Long) As Double
     ValueMatrix = CleanValue(fg.TextMatrix(Row, Col))
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=fg,fg,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets whether text within a cell should be allowed to wrap."
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = "Style"
    WordWrap = fg.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    fg.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property







