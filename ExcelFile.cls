VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelFileV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' Class om versie 2 Excel file uit te schrijven
' ondersteund: strings, integers, doubles
' 1 sheet per excel file
' gebaseerd op een stuk code gevonden op het internet
' Arnout: 12 oktober 1999

'Beginning Of File record
Private Type BOF
  opcode As Integer
  Length As Integer
  version As Integer
  ftype As Integer
End Type

'End Of File record
Private Type EOF
  opcode As Integer
  Length As Integer
End Type

'Integer record
Private Type tInteger
  opcode As Integer
  Length As Integer
  Row As Long
  Col As Long
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  Value As Integer
End Type

'Number = double record
Private Type tNumber
  opcode As Integer
  Length As Integer
  Row As Long
  Col As Long
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  Value As Double
End Type

'Label (Text) record
Private Type tLabel
  opcode As Integer
  Length As Integer
  Row As Integer
  Col As Integer
  rgbAttr1 As Byte
  rgbAttr2 As Byte
  rgbAttr3 As Byte
  strLength As Byte
End Type

Dim fhFile As Integer
Dim bof1 As BOF
Dim eof1 As EOF
Dim l1 As tLabel
Dim i1 As tInteger
Dim n1 As tNumber

Private Sub Class_Initialize()
  'Set up default values for records
  'These should be the values that are the same for every record

  With bof1
    .opcode = 9
    .Length = 4
    .version = 2
    .ftype = 10
  End With

  With eof1
    .opcode = 10
  End With

  With l1
    .opcode = 4
    .Length = 10
    .Row = 0
    .Col = 0
    .rgbAttr1 = 0
    .rgbAttr2 = 0
    .rgbAttr3 = 0
    .strLength = 2
  End With

  With i1
    .opcode = 2
    .Length = 9
    .Row = 0
    .Col = 0
    .rgbAttr1 = 0
    .rgbAttr2 = 0
    .rgbAttr3 = 0
    .Value = 0
  End With

  With n1
    .opcode = 3
    .Length = 15
    .Row = 0
    .Col = 0
    .rgbAttr1 = 0
    .rgbAttr2 = 0
    .rgbAttr3 = 0
    .Value = 0
  End With

End Sub

Public Sub OpenFile(ByVal FileName As String)
  fhFile = FreeFile
  Open FileName For Binary As #fhFile
  Put #fhFile, , bof1
End Sub

Public Sub CloseFile()
  Put #fhFile, , eof1
  Close #fhFile
End Sub

Function EWriteString(ExcelRow As Long, ExcelCol As Long, Text As String)
  Dim b As Byte, l As Byte, a As Byte
'Length of the text portion of the record
  l = Len(Text)
  l1.strLength = l

'Total length of the record
  l1.Length = 8 + l1.strLength

  l1.Row = ExcelRow - 1
  l1.Col = ExcelCol - 1

'Put record header
  Put #fhFile, , l1

'Then the actual string data
'We have to write the string one character at a time, so we loop
'through all characters in the string, assign thier ASCII value to b
'and do a Put on b (which is declared as Byte)
  For a = 1 To l
    b = Asc(Mid$(Text, a, 1))
    Put #fhFile, , b
  Next

End Function

Function EWriteInteger(ExcelRow As Long, ExcelCol As Long, Value As Integer)
  With i1
    .Row = ExcelRow - 1
    .Col = ExcelCol - 1
    .Value = Value
  End With

  Put #fhFile, , i1
End Function

Function EWriteDouble(ExcelRow As Long, ExcelCol As Long, Value As Double)
  With n1
    .Row = ExcelRow - 1
    .Col = ExcelCol - 1
    .Value = Value
  End With

  Put #fhFile, , n1
End Function

