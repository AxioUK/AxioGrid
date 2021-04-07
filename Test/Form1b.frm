VERSION 5.00
Object = "{07242263-ECFB-49C3-84AD-1CD9A3F8DB91}#10.0#0"; "AXGRIDKM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AxGrid/AxBiGrid V2.0 TEST FORM"
   ClientHeight    =   7770
   ClientLeft      =   4950
   ClientTop       =   2730
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   15810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "AxGrid"
      Height          =   7380
      Left            =   45
      TabIndex        =   37
      Top             =   315
      Width           =   6825
      Begin AxioGrid.AxGrid AxGrid1 
         Height          =   3360
         Left            =   135
         TabIndex        =   76
         Top             =   285
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   5927
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLinesFixed  =   2
         Appearance      =   0
         BackColorFixed  =   -2147483626
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColorFixed  =   8421504
         MouseIcon       =   "Form1b.frx":0000
         Rows            =   10
         ScrollTrack     =   0   'False
      End
      Begin VB.CommandButton cmdCreateFullTable 
         Caption         =   "CREATE FULLTABLE"
         Height          =   270
         Index           =   2
         Left            =   4830
         TabIndex        =   75
         Top             =   5475
         Width           =   1830
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Check Column"
         Height          =   330
         Left            =   4710
         TabIndex        =   74
         Top             =   3795
         Width           =   1230
      End
      Begin VB.CommandButton cmdExportXLS 
         Caption         =   "Exportar XLS"
         Height          =   360
         Index           =   0
         Left            =   2280
         TabIndex        =   73
         Top             =   6915
         Width           =   1275
      End
      Begin VB.CheckBox chkAutoNum 
         Caption         =   "Autonumerate First Col"
         Height          =   240
         Left            =   2160
         TabIndex        =   63
         Top             =   3855
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change 2nd Row BackColor"
         Height          =   510
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   5205
         Width           =   1980
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change 3nd Row CellAlignment"
         Height          =   510
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   5715
         Width           =   1980
      End
      Begin VB.CommandButton cmdAutosize 
         Caption         =   "Autosize 3 first Columns"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   6240
         Width           =   1980
      End
      Begin VB.CommandButton cmdSum 
         Caption         =   "Sumar Matrix"
         Height          =   630
         Index           =   0
         Left            =   5805
         TabIndex        =   59
         Top             =   4440
         Width           =   915
      End
      Begin VB.CommandButton cmdCreateTable 
         Caption         =   "CREATE TABLE"
         Height          =   270
         Index           =   0
         Left            =   4830
         TabIndex        =   58
         Top             =   5205
         Width           =   1830
      End
      Begin VB.TextBox txtCOL 
         Height          =   315
         Left            =   1440
         TabIndex        =   57
         Top             =   4455
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   4605
         TabIndex        =   56
         Top             =   4455
         Width           =   1170
      End
      Begin VB.CommandButton cmdLoadDB2 
         Caption         =   "LoadfromDB"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   6930
         Width           =   1080
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4605
         TabIndex        =   54
         Top             =   4785
         Width           =   1170
      End
      Begin VB.TextBox txtRow 
         Height          =   285
         Left            =   1440
         TabIndex        =   53
         Top             =   4785
         Width           =   720
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "INSERT"
         Height          =   270
         Index           =   0
         Left            =   4830
         TabIndex        =   52
         Top             =   6285
         Width           =   1830
      End
      Begin VB.CommandButton cmdDrpTable 
         Caption         =   "DROP TABLE"
         Height          =   270
         Index           =   0
         Left            =   4830
         TabIndex        =   51
         Top             =   5745
         Width           =   1830
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "SELECT"
         Height          =   270
         Index           =   0
         Left            =   4830
         TabIndex        =   50
         Top             =   6015
         Width           =   1830
      End
      Begin VB.CommandButton cmdLoadDB2 
         Caption         =   "with Index"
         Height          =   345
         Index           =   1
         Left            =   1215
         TabIndex        =   49
         Top             =   6930
         Width           =   885
      End
      Begin VB.CommandButton cmdInsrtSel 
         Caption         =   "INSERT SELECT"
         Height          =   270
         Index           =   0
         Left            =   4830
         TabIndex        =   48
         Top             =   6555
         Width           =   1830
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load to Object"
         Height          =   345
         Index           =   0
         Left            =   2295
         TabIndex        =   47
         Top             =   6570
         Width           =   1260
      End
      Begin VB.CheckBox chkMove 
         Caption         =   "On Enter Key Move Down"
         Height          =   270
         Left            =   2160
         TabIndex        =   46
         Top             =   4080
         Width           =   2415
      End
      Begin VB.CheckBox chkEditable 
         Caption         =   "Editable"
         Height          =   210
         Left            =   270
         TabIndex        =   45
         Top             =   3825
         Width           =   1635
      End
      Begin VB.CheckBox chkAltColor 
         Caption         =   "Alternate Color"
         Height          =   210
         Left            =   255
         TabIndex        =   44
         Top             =   4080
         Width           =   1635
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Controls Columnas 1-2-3"
         Height          =   360
         Left            =   2310
         TabIndex        =   43
         Top             =   5220
         Width           =   2415
      End
      Begin VB.CommandButton cmdLoadRST 
         Caption         =   "Load From RST"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   6600
         Width           =   1980
      End
      Begin VB.CommandButton cmdShowInfoBar 
         Caption         =   "InfoBar"
         Height          =   360
         Index           =   0
         Left            =   2310
         TabIndex        =   41
         Top             =   5595
         Width           =   1155
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Flat Style"
         Height          =   360
         Index           =   1
         Left            =   2310
         TabIndex        =   40
         Top             =   5970
         Width           =   1155
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3495
         TabIndex        =   39
         Top             =   6225
         Width           =   1260
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   3495
         TabIndex        =   38
         Top             =   5610
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumar Columna"
         Height          =   195
         Left            =   210
         TabIndex        =   67
         Top             =   4500
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado CalculateColumn"
         Height          =   195
         Left            =   2550
         TabIndex        =   66
         Top             =   4500
         Width           =   1950
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado CalculateRow"
         Height          =   195
         Left            =   2550
         TabIndex        =   65
         Top             =   4875
         Width           =   1950
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplicar Fila"
         Height          =   195
         Left            =   210
         TabIndex        =   64
         Top             =   4875
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Fixed Rows +1"
      Height          =   300
      Left            =   6060
      TabIndex        =   34
      Top             =   0
      Width           =   1710
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Unload Me"
      Height          =   300
      Left            =   7815
      TabIndex        =   33
      Top             =   0
      Width           =   1710
   End
   Begin VB.CommandButton Command12 
      Caption         =   "To Advanced Test"
      Height          =   300
      Left            =   4305
      TabIndex        =   32
      Top             =   0
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AxBiGrid"
      Height          =   7395
      Left            =   6900
      TabIndex        =   0
      Top             =   315
      Width           =   8835
      Begin AxioGrid.AxBiGrid AxBiGrid1 
         Height          =   3645
         Left            =   180
         TabIndex        =   77
         Top             =   270
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   6429
         SplitterPos     =   1985
         EnterKeyBehaviour=   0
         GridLinesFixed  =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1b.frx":001C
         RowHeightMin    =   255
         Rows            =   2
         ScrollTrack     =   0   'False
         GridLinesFixed  =   2
         Appearance      =   0
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1b.frx":0038
         RowHeightMin    =   255
         Rows            =   2
         ScrollTrack     =   0   'False
      End
      Begin VB.Frame Frame3 
         Caption         =   "Load..."
         Height          =   1860
         Left            =   5085
         TabIndex        =   68
         Top             =   5280
         Width           =   1710
         Begin VB.CommandButton cmdLoadDB2 
            Caption         =   "...to Controls"
            Height          =   330
            Index           =   3
            Left            =   180
            TabIndex        =   72
            Top             =   1410
            Width           =   1365
         End
         Begin VB.CommandButton cmdLoadDB2 
            Caption         =   "... from DB"
            Height          =   330
            Index           =   2
            Left            =   180
            TabIndex        =   71
            Top             =   1020
            Width           =   1365
         End
         Begin VB.CommandButton cmdLoadRST 
            Caption         =   "...RST to Right"
            Height          =   330
            Index           =   1
            Left            =   180
            TabIndex        =   70
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton cmdLoadToRow 
            Caption         =   "...RST to Left"
            Height          =   330
            Left            =   180
            TabIndex        =   69
            Top             =   630
            Width           =   1365
         End
      End
      Begin VB.CommandButton cmdExportXLS 
         Caption         =   "Exportar XLS"
         Height          =   360
         Index           =   1
         Left            =   195
         TabIndex        =   36
         Top             =   6915
         Width           =   1575
      End
      Begin VB.CommandButton cmdShowInfoBar 
         Caption         =   "InfoBar"
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   35
         Top             =   5295
         Width           =   1350
      End
      Begin VB.TextBox txtinfoBar 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Text            =   "Text7"
         Top             =   6435
         Width           =   1350
      End
      Begin VB.ListBox lstTextGrid 
         Height          =   840
         Left            =   1920
         TabIndex        =   30
         Top             =   5595
         Width           =   1350
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Flat Style"
         Height          =   360
         Index           =   0
         Left            =   195
         TabIndex        =   29
         Top             =   6540
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Fix Splitter"
         Height          =   300
         Left            =   4800
         TabIndex        =   28
         Top             =   4785
         Width           =   1215
      End
      Begin VB.TextBox txtSplitPos 
         Height          =   285
         Left            =   4050
         TabIndex        =   26
         Top             =   4785
         Width           =   720
      End
      Begin VB.CommandButton Command9 
         Caption         =   "AddItem Right"
         Height          =   495
         Index           =   1
         Left            =   990
         TabIndex        =   25
         Top             =   6030
         Width           =   780
      End
      Begin VB.CommandButton Command9 
         Caption         =   "AddItem Left"
         Height          =   495
         Index           =   0
         Left            =   195
         TabIndex        =   24
         Top             =   6030
         Width           =   780
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SelectionMode Free"
         Height          =   270
         Left            =   195
         TabIndex        =   23
         Top             =   4830
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Controls Cols 1-2"
         Height          =   360
         Left            =   195
         TabIndex        =   22
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   360
         Left            =   195
         TabIndex        =   21
         Top             =   5655
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4050
         TabIndex        =   16
         Top             =   4080
         Width           =   720
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6240
         TabIndex        =   15
         Top             =   4080
         Width           =   1170
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6240
         TabIndex        =   14
         Top             =   4410
         Width           =   1170
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4050
         TabIndex        =   13
         Top             =   4410
         Width           =   720
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change 2nd Row BackColor"
         Height          =   600
         Index           =   1
         Left            =   3405
         TabIndex        =   12
         Top             =   5295
         Width           =   1545
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change 3nd Row CellAlignment"
         Height          =   600
         Index           =   1
         Left            =   3405
         TabIndex        =   11
         Top             =   5910
         Width           =   1545
      End
      Begin VB.CommandButton cmdAutosize 
         Caption         =   "Autosize 3 first Columns"
         Height          =   600
         Index           =   1
         Left            =   3405
         TabIndex        =   10
         Top             =   6525
         Width           =   1545
      End
      Begin VB.CommandButton cmdSum 
         Caption         =   "Sumar Matrix"
         Height          =   645
         Index           =   1
         Left            =   7485
         TabIndex        =   9
         Top             =   4080
         Width           =   750
      End
      Begin VB.CommandButton cmdCreateTable 
         Caption         =   "CREATE TABLE"
         Height          =   300
         Index           =   1
         Left            =   6945
         TabIndex        =   8
         Top             =   5280
         Width           =   1665
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "INSERT"
         Height          =   300
         Index           =   1
         Left            =   6945
         TabIndex        =   7
         Top             =   6330
         Width           =   1665
      End
      Begin VB.CommandButton cmdDrpTable 
         Caption         =   "DROP TABLE"
         Height          =   300
         Index           =   1
         Left            =   6945
         TabIndex        =   6
         Top             =   5625
         Width           =   1665
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "SELECT"
         Height          =   300
         Index           =   1
         Left            =   6945
         TabIndex        =   5
         Top             =   5970
         Width           =   1665
      End
      Begin VB.CommandButton cmdInsrtSel 
         Caption         =   "INSERT SELECT"
         Height          =   300
         Index           =   1
         Left            =   6945
         TabIndex        =   4
         Top             =   6675
         Width           =   1665
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alternate Color"
         Height          =   210
         Left            =   195
         TabIndex        =   3
         Top             =   4330
         Width           =   1635
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Editable"
         Height          =   210
         Left            =   195
         TabIndex        =   2
         Top             =   4110
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Enter Key Move Down"
         Height          =   270
         Left            =   195
         TabIndex        =   1
         Top             =   4550
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Splitter Position"
         Height          =   195
         Left            =   2910
         TabIndex        =   27
         Top             =   4830
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumar Columna"
         Height          =   195
         Left            =   2820
         TabIndex        =   20
         Top             =   4125
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CalculateColumn"
         Height          =   195
         Left            =   4965
         TabIndex        =   19
         Top             =   4125
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CalculateRow"
         Height          =   195
         Left            =   5160
         TabIndex        =   18
         Top             =   4500
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplicar Fila"
         Height          =   195
         Left            =   2910
         TabIndex        =   17
         Top             =   4455
         Width           =   1110
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I     As Integer
Dim Conn  As String
Dim Rst   As New ADODB.Recordset
Dim Cnn   As New ADODB.Connection


Private Sub AxBiGrid1_CellTextChange(xGrid As AxioGrid.eSideGrid, ByVal Row As Long, ByVal Col As Long)
If AxBiGrid1.ColObject(eLeftGrid, 2) = eTextBoxColumn Then
  AxBiGrid1.SetColObject(eLeftGrid, 2, True) = eComboBoxColumn
End If

End Sub


Private Sub AxBiGrid1_Click(xGrid As AxioGrid.eSideGrid, Row As Long, Col As Long)
If lstTextGrid.ListIndex = 3 Then AxBiGrid1.InfoBarText = txtinfoBar.Text & ":" & AxBiGrid1.Text(eLeftGrid) & ":" & AxBiGrid1.Text(eRightGrid)

End Sub

Private Sub AxBiGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
txtSplitPos.Text = AxBiGrid1.SplitterPos
End Sub

Private Sub AxGrid1_ButtonClick(ByVal Row As Long, ByVal Col As Long)
MsgBox "Lin:" & Row & "/Col:" & Col & " - " & AxGrid1.TextMatrix(Row, Col)
End Sub


Private Sub AxGrid1_Click(lRow As Long, lCol As Long)
With AxGrid1
   If List1.ListIndex = 3 And .Col <> 5 Then
      .InfoBarText = Text7.Text & ":" & AxGrid1.Text
   End If
   
End With

End Sub

Private Sub AxGrid1_ListClick(ByVal Row As Long, ByVal Col As Long, ByVal iListIndex As Long)
MsgBox Row & " : " & Col
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
       AxBiGrid1.EnterKeyBehaviour = axEKMoveDown
    Else
       AxBiGrid1.EnterKeyBehaviour = axEKMoveRight
    End If

End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      AxBiGrid1.Editable = True
   Else
      AxBiGrid1.Editable = False
   End If

End Sub

Private Sub Check3_Click()
   If Check3.Value = vbChecked Then
      AxBiGrid1.BackColorAlternate = &HFFFFC0
   Else
      AxBiGrid1.BackColorAlternate = vbWhite
   End If
End Sub

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
  AxBiGrid1.SelectionMode = flexSelectionFree
Else
  AxBiGrid1.SelectionMode = flexSelectionByRow
End If
End Sub

Private Sub chkAltColor_Click()
   If chkAltColor.Value = vbChecked Then
      AxGrid1.BackColorAlternate = &HC0E0FF
   Else
      AxGrid1.BackColorAlternate = vbWhite
   End If
End Sub

Private Sub chkEditable_Click()
   If chkEditable.Value = vbChecked Then
      AxGrid1.Editable = True
   Else
      AxGrid1.Editable = False
   End If
End Sub

Private Sub chkMove_Click()
    If chkMove.Value = vbChecked Then
       AxGrid1.EnterKeyBehaviour = axEKMoveDown
    Else
       AxGrid1.EnterKeyBehaviour = axEKMoveRight
    End If
End Sub

Private Sub chkAutoNum_Click()
AxGrid1.AutoNumFirstCol = Not AxGrid1.AutoNumFirstCol
End Sub

Private Sub cmdCheck_Click()
Dim sString As String

AxGrid1.CellCheckValue(1, 5) = True
AxGrid1.CellCheckValue(2, 5) = True
AxGrid1.CellCheckValue(3, 5) = True

For I = 1 To AxGrid1.Rows - 1
   sString = sString & AxGrid1.CellCheckValue(I, 5) & vbNewLine
Next I

MsgBox sString
End Sub

Private Sub cmdCreateFullTable_Click(Index As Integer)
With AxGrid1
  .ADOConnection = Conn
  I = I + 1
  If .ADOCreateGridTable("FTabla" & I, True) = True Then
    MsgBox "Full Tabla Creada!"
  Else
    MsgBox "Error al Crear Full Tabla!"
  End If
End With

End Sub

Private Sub cmdDrpTable_Click(Index As Integer)
AxGrid1.ADODropTable "MiTabla"
End Sub

Private Sub cmdExportXLS_Click(Index As Integer)
Select Case Index
  Case Is = 0
    AxGrid1.SaveAsExcel App.Path & "\TestAxGrid_Export.xls", "AxGRID1", False
  Case Is = 1
    AxBiGrid1.SaveAsExcel App.Path & "\TestBiGrid_Export.xls"
End Select

End Sub

Private Sub cmdInsert_Click(Index As Integer)
'Dim sFields As String, I As Integer
'
'For I = 1 To AxGrid1.Cols - 1
'  sFields = sFields & AxGrid1.TextMatrix(0, I) & ","
'Next I
'
'If Right$(sFields, 1) = "," Then
'  sFields = Mid$(sFields, 1, Len(sFields) - 1)
'End If

With AxGrid1
  '.ADOTable = "MiTabla"
  '.ADOFields = sFields
  .ADOInsertRow AxGrid1.Row, 0, 2
End With

End Sub

Private Sub cmdLoadRST_Click(Index As Integer)
Dim Consulta As String

Consulta = "SELECT * FROM Unidad"

Cnn.Open Conn
Rst.Open Consulta, Cnn, adOpenDynamic, adLockOptimistic

Select Case Index
   Case 0
    AxGrid1.LoadfromRST Rst, True, True
   
   Case 1
   
    LoadSideGrid Rst, AxBiGrid1, eLeftGrid
    
End Select

Rst.Close
Cnn.Close

End Sub

Public Sub LoadSideGrid(sqlRst As ADODB.Recordset, FlexGrid As AxBigrid, eSideG As eSideGrid)
On Local Error GoTo error_handler
    Dim Columna As Integer, Fila As Integer

With FlexGrid
  ' -- limpiamos el grid
  .ClearGrid BothGrids
  ' -- Deshabilito el repintado del control para acelerar carga
  .Redraw = False
  .Rows = 2 ' --> Cantidad inicial de filas: Evita Error al Reutilizar Grid...
  ' -- Modo de encabezados
  .FixedRows = 1
  .FixedCols = 0
  ' -- Cantidad de filas y columnas
  .Rows = 1
  If eSideG = eLeftGrid Then
    .ColsLeft = sqlRst.Fields.Count
  Else
    .ColsRight = sqlRst.Fields.Count
  End If
  
  ' -- Recorrer los campos del recordset
  For Columna = 0 To sqlRst.Fields.Count - 1
    ' -- Añade el título del campo al encabezado de columna
    .TextMatrix(eSideG, 0, Columna) = sqlRst.Fields(Columna).Name
  Next Columna
  Fila = 1
  ' -- Recorrer todos los registros del recordset
  Do While Not sqlRst.EOF
    .Rows = .Rows + 1 ' Añade una nueva fila
    For Columna = 0 To sqlRst.Fields.Count - 1
      ' -- Combobar que el valor no es nulo
      If Not IsNull(sqlRst.Fields(Columna).Value) Then
        ' -- Agrega el registro en la fila y columna específica
        .TextMatrix(eSideG, Fila, Columna) = sqlRst.Fields(Columna).Value
      End If
    Next
    ' -- Siguiente registro
    sqlRst.MoveNext
    Fila = Fila + 1 'Incrementa la fila
  Loop
  ' -- Cierra el recordset y la conexión abierta
  If sqlRst.State = adStateOpen Then sqlRst.Close
  
  ' -- Volver a Habilitar el repintado del Grid
  .Redraw = True
End With
'WriteLog "Mostrando: " & FlexGrid.Rows - 1 & " registros encontrados"
' -- Error
Exit Sub
error_handler:
If sqlRst.State = adStateOpen Then sqlRst.Close
'WriteLog "LoadFlex.Error: " & Err.Number & ":" & Err.Description
'WriteLog Err.Number & ":" & Err.Description & vbCrLf & "No existen datos en Base de Datos" & sForm.Tag
End Sub

Private Sub cmdLoadToRow_Click()
Dim SQLString As String
Dim iRow As Integer, Columna As Integer

SQLString = "SELECT EQ.EQID, EQ.SERIE FROM EQUIPOS;"
            
  Cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Programming Folder\SERVTEC V2.0\DBMANT.accdb;Persist Security Info=False;"
  Rst.Open SQLString, Cnn, adOpenDynamic, adLockOptimistic
  LoadSideGrid Rst, AxBiGrid1, eLeftGrid
  Rst.Close
  
'For iRow = 1 To AxBiGrid1.Rows - 1
'  SQLString = "SELECT CAL.FECHACAL FROM CALIBRED AS CAL WHERE CAL.EQUIPO=" & AxBiGrid1.TextMatrix(eLeftGrid, iRow, 0) & ";"
'  Rst.Open SQLString, Cnn, adOpenDynamic, adLockOptimistic
'  'AxBiGrid1.LoadtoRow iRow, Rst
'  For Columna = 0 To Rst.RecordCount - 1
'    MsgBox Rst.RecordCount
'    AxBiGrid1.TextMatrix(eRightGrid, iRow, Columna) = Rst(Columna)
'  Next Columna
'  If Rst.State = adStateOpen Then Rst.Close
'Next
  
  Cnn.Close
  
End Sub

Private Sub cmdShowInfoBar_Click(Index As Integer)
If Index = 1 Then
  AxBiGrid1.ShowInfoBar = Not AxBiGrid1.ShowInfoBar
Else
  AxGrid1.ShowInfoBar = Not AxGrid1.ShowInfoBar
End If
End Sub

Private Sub cmdSum_Click(Index As Integer)
If Index = 0 Then
    Text1.Text = AxGrid1.CalculateMatrix(axSTSum, 1, 2, 8, 2)
Else
    Text5.Text = AxBiGrid1.CalculateColumn(eRightGrid, axSTSum, 2, 1, 10)
End If
End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
     AxGrid1.cell(axcpCellBackColor, 2, 1, 2, AxGrid1.Cols - 1) = vbBlue
Else
    AxBiGrid1.cell(axcpCellBackColor, 2, 0, 2, AxBiGrid1.ColsRight - 1, eRightGrid) = vbBlue
End If
End Sub

Private Sub Command10_Click()
AxBiGrid1.SplitterFixed = Not AxBiGrid1.SplitterFixed
If Command10.Caption = "Fix Splitter" Then
  Command10.Caption = "UnFix Splitter"
Else
  Command10.Caption = "Fix Splitter"
End If
End Sub

Private Sub Command11_Click(Index As Integer)
If Index = 0 Then
  If AxBiGrid1.Appearance = flex3D Then
    AxBiGrid1.Appearance = flexFlat
    Command11(0).Caption = "3D Style"
  Else
    AxBiGrid1.Appearance = flex3D
    Command11(0).Caption = "Flat Style"
  End If
Else
  If AxGrid1.Appearance = flex3D Then
    AxGrid1.Appearance = flexFlat
    Command11(1).Caption = "3D Style"
  Else
    AxGrid1.Appearance = flex3D
    Command11(1).Caption = "Flat Style"
  End If
End If
End Sub

Private Sub Command13_Click()
Unload Me
End Sub

Private Sub Command14_Click()
AxGrid1.FixedRows = AxGrid1.FixedRows + 1
AxBiGrid1.FixedRows = AxBiGrid1.FixedRows + 1
End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then
     AxGrid1.cell(axcpCellAlignment, 3, 1, 3, AxGrid1.Cols - 1) = 3
Else
    AxBiGrid1.cell(axcpCellAlignment, 3, 1, 3, AxBiGrid1.ColsRight - 1, eRightGrid) = 3
End If
End Sub

Private Sub cmdAutosize_Click(Index As Integer)
If Index = 0 Then
    AxGrid1.AutoSizeMode = axAutoSizeColWidth
    AxGrid1.AutoSizeCols 0, 2

Else
    AxBiGrid1.AutoSizeMode = axAutoSizeColWidth
    AxBiGrid1.AutoSizeCols eRightGrid, 0, 2
End If
End Sub

Private Sub cmdCreateTable_Click(Index As Integer)

With AxGrid1
  .ADOConnection = Conn
  If .ADOCreateTable("MiTabla", 1, 2) = True Then
    MsgBox "Tabla Creada!"
  Else
    MsgBox "Error al Crear Tabla!"
  End If
End With

End Sub


Private Sub Command3_Click()
With AxBiGrid1
  .SetColObject(eLeftGrid, 1, False) = eTextBoxColumn
  .SetColObject(eLeftGrid, 2, False) = eButtonColumn
  .SetColObject(eRightGrid, 1, False) = eComboBoxColumn
  .SetColObject(eRightGrid, 2, False) = eListBoxColumn
  For I = 0 To 10
    .AddItemObject oComboBox, "Item_" & I
    .AddItemObject oListBox, "Value_" & I + 100
  Next I
End With

End Sub

Private Sub Command4_Click()
With AxBiGrid1
  .ClearGrid BothGrids
  .Rows = 3
  .SetColObject(eLeftGrid, 2, False) = eTextBoxColumn
  .AddRowsOnDemand = Not .AddRowsOnDemand
End With
End Sub

Private Sub Command8_Click()
Dim I As Integer
For I = 1 To 10
  AxGrid1.AddItemObject "Item " & I
Next I

With AxGrid1
   .Cols = 6
   .SetColObject(1) = eTextBoxColumn
   .SetColObject(2) = eButtonColumn
   .SetColObject(3) = eComboBoxColumn
   .SetColObject(4) = eListBoxColumn
   .SetColObject(5) = eCheckBoxColumn
End With
End Sub

Private Sub Command9_Click(Index As Integer)
Select Case Index
   Case 0
    AxBiGrid1.Row(eLeftGrid) = 2
    AxBiGrid1.AddItem eLeftGrid, vbTab & "TEST ADDITEM L", 2
   Case 1
    AxBiGrid1.Row(eRightGrid) = 2
    AxBiGrid1.AddItem eRightGrid, "TEST ADDITEM R", 2
End Select
End Sub

Private Sub Form_Load()
Dim I As Integer
With AxGrid1
  .TextMatrix(0, 1) = "Name"
  .TextMatrix(0, 2) = "Salary"
  .ColDisplayFormat(2) = "#0.00"
  .Cols = 4
  .Rows = 15
  .ColWidth(0) = 600
    
  For I = 1 To .Rows - 1
    .TextMatrix(I, 1) = "ABC_" & I
    .TextMatrix(I, 2) = I * 3.05
    .ColDisplayFormat(3) = "$#,###.#0"
    .TextMatrix(I, 3) = I * 12.05
  Next I
    
  'Implemented only for Numeric Entry
  .ColInputMask(2) = "000.00"
  .TextMatrix(9, 0) = "Total"
  
End With
   
With AxBiGrid1
  .ColsLeft = 3
  .ColsRight = 5
  .Rows = 15
  .TextMatrix(eLeftGrid, 0, 1) = "PEOPLE"
  .TextMatrix(eRightGrid, 0, 1) = "STATE"
  .TextMatrix(eRightGrid, 0, 2) = "SALARY"
  .ColDisplayFormat(2) = "#0.00"
  .ColWidth(BothGrids, 0) = 500
  .ColWidth(LeftGrid, 1) = 1800
  .ColWidth(RightGrid, 1) = 1800
  
  For I = 1 To .Rows - 1
    .TextMatrix(eLeftGrid, I, 1) = "NOMBRE_" & I
    .TextMatrix(eLeftGrid, I, 2) = I * 103.5
    .TextMatrix(eRightGrid, I, 1) = "ABC_" & I
    .TextMatrix(eRightGrid, I, 2) = I * 3.05
  Next I
    
  'Implemented only for Numeric Entry
  .ColInputMask(2) = "000.00"
  .TextMatrix(eRightGrid, 9, 0) = "Total"
  
End With

With lstTextGrid
    .AddItem "BothGrids"
    .AddItem "LeftGrids"
    .AddItem "RightGrids"
    .AddItem "Custom"
End With

With List1
    .AddItem "CellGridInfo"
    .AddItem "RowGridInfo"
    .AddItem "ColGridInfo"
    .AddItem "CustomText"
End With

End Sub

Private Sub AxGrid1_BeforeEdit(Cancel As Boolean)
     If AxGrid1.Col = 1 Then
        If chkAutoNum.Value = vbUnchecked Then
            Cancel = True
        End If
     End If
End Sub

Private Sub List1_Click()
AxGrid1.SetInfoBar = List1.ListIndex
End Sub

Private Sub lstTextGrid_Click()
AxBiGrid1.SetInfoBar = lstTextGrid.ListIndex
End Sub

Private Sub txtCOL_Change()
On Error Resume Next
Text1.Text = AxGrid1.CalculateColumn(axSTSum, txtCOL.Text, 1, AxGrid1.Rows - 1)
End Sub

Private Sub txtRow_Change()
On Error Resume Next
Text2.Text = AxGrid1.CalculateRow(axSTMultiply, txtRow.Text, 2, 3)
End Sub

Private Sub txtSplitPos_Change()
AxBiGrid1.SplitterPos = txtSplitPos.Text
End Sub
