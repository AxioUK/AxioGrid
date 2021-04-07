VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type FgCol
    ColDisplayFormat As String
    ColInputMask As String
    ColType As eColumnType
    ColLocked As eBoolean
End Type

Public Enum eBoolean
    isFalse = 0
    isTrue = 1
End Enum

Public Enum eCellProperty
    axcpCellAlignment
    axcpCellFontName
    axcpCellFontSize
    axcpCellFontBold
    axcpCellForeColor
    axcpCellBackColor
End Enum

Public Enum eAlignCols
    'Constant Value Description
    eAlignLeftTop = 0       'The column content is aligned left, top.
    eAlignLeftCenter = 1    'Default for strings. The column content is aligned left, center.
    eAlignLeftBottom = 2    'The column content is aligned left, bottom.
    eAlignCenterTop = 3     'The column content is aligned center, top.
    eAlignCenterCenter = 4  'The column content is aligned center, center.
    eAlignCenterBottom = 5  'The column content is aligned center, bottom.
    eAlignRightTop = 6      'The column content is aligned right, top.
    eAlignRightCenter = 7   'Default for numbers. The column content is aligned right, center.
    eAlignRightBottom = 8   'The column content is aligned right, bottom.
    eAlignGeneral = 9       'The column content is of general alignment. This is "left, center" for strings and "right, center" for numbers.
End Enum

Public Enum eAutoSizeSetting
    axAutoSizeColWidth
    axAutoSizeRowHeight
End Enum

Public Enum eSubTotalSettings
    axSTMultiply
    axSTCount
    axSTMax
    axSTMin
    axSTSum
End Enum

Public Enum eEnterkeyBehaviour
    axEKMoveDown
    axEKMoveRight
    axEKNone
End Enum

Public Enum eColumnType
    eRemoveObject = 0
    eComboBoxColumn = 5
    eListBoxColumn = 6
    eTextBoxColumn = 7
    eButtonColumn = 8
    eCheckBoxColumn = 9
End Enum

Public Enum eSideGrid
    eLeftGrid = 1
    eRightGrid = 2
End Enum

Public Enum eSideGrid2
    BothGrids = 0
    LeftGrid = 1
    RightGrid = 2
End Enum

Public Enum eTypeInfoBar
    BothGridsInfo = 0
    LeftGridInfo = 1
    RightGridInfo = 2
    CustomInfo = 3
End Enum

Public Enum eTypeSingleInfoBar
    CellGridInfo = 0
    RowGridInfo = 1
    ColGridInfo = 2
    CustomText = 3
End Enum

Public Enum eTypeControl
    oComboBox = 5
    oListBox = 6
End Enum

Public Enum eInputChar
    [AllAllowed] = 0
    [IntegerOnly] = 1
    [LettersOnly] = 2
    [DecimalNumbers] = 3
    [UpCaseLetterOnly] = 4
    [LoCaseLetterOnly] = 5
End Enum

'ENUMS----------------------
Public Enum RegionalConstant
  LOCALE_SCURRENCY = &H14
  LOCALE_SCOUNTRY = &H6
  LOCALE_SDATE = &H1D
  LOCALE_SDECIMAL = &HE
  LOCALE_SLANGUAGE = &H2
  LOCALE_SLONGDATE = &H20
  LOCALE_SMONDECIMALSEP = &H16
  LOCALE_SMONGROUPING = &H18
  LOCALE_SMONTHOUSANDSEP = &H17
  LOCALE_SNATIVECTRYNAME = &H8
  LOCALE_SNATIVECURRNAME = &H1008
  LOCALE_SNATIVEDIGITS = &H13
  LOCALE_SNEGATIVESIGN = &H51
  LOCALE_SSHORTDATE = &H1F
  LOCALE_STIME = &H1E
  LOCALE_STIMEFORMAT = &H1003
End Enum


