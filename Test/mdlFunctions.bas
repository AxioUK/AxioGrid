Attribute VB_Name = "mdlFunctions"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : FormatoRUT
' Author    : AxioUK
' Date      : 05-11-2011
'---------------------------------------------------------------------------------------
Public Function FormatoRUT(sRUT As String) As String
Dim strRut As String
   If Len(sRUT) = 10 Then
      strRut = Mid(sRUT, 1, 2) & "." & Mid(sRUT, 3, 3) & "." & Mid(sRUT, 6, 5)
   ElseIf Len(sRUT) = 9 Then
      strRut = Mid(sRUT, 1, 1) & "." & Mid(sRUT, 2, 3) & "." & Mid(sRUT, 5, 5)
   Else
      strRut = sRUT
   End If
   FormatoRUT = strRut
End Function

'---------------------------------------------------------------------------------------
' Procedure : EsRUT
' Author    : AxioUK
' Date      : 05-11-2011
'---------------------------------------------------------------------------------------
Public Function EsRUT(CadenA As String) As Boolean
Dim I As Byte, Z As Byte
Dim CadenaLimpiA As String
Dim DiG As String, XXXX As Byte

If CadenA <> Empty And Val(CadenA) <> 0 Then
    'Limpia Cadena
    For I = 1 To Len(CadenA)
        If (Mid(CadenA, I, 1)) = "-" Or (Mid(CadenA, I, 1)) = "." Then
            'pasa al siguiente espacio
        Else
            CadenaLimpiA = CadenaLimpiA + Mid(CadenA, I, 1)
        End If
    Next
    
    'Prepara Variables
    CadenA = CadenaLimpiA
    DiG = (Mid(CadenaLimpiA, (Len(CadenaLimpiA)), 1))
    If Asc(DiG) <= 47 Or Asc(DiG) >= 58 Then
        If DiG = "K" Or DiG = "k" Then
            DiG = "10"
        Else
           DiG = "12"
        End If
    End If
    
    CadenaLimpiA = Empty
    
    For I = 1 To (Len(CadenA) - 1)
        CadenaLimpiA = CadenaLimpiA + (Mid(CadenA, I, 1))
    Next
    
    CadenA = Empty
    I = Empty
    I = (Len(CadenaLimpiA))
    Z = 2
    While I <> 0
        If Z <> 8 Then
            CadenA = Val(CadenA) + (Val((Mid(CadenaLimpiA, I, 1))) * Z)
            Z = Z + 1
        Else
            Z = 2
            CadenA = Val(CadenA) + (Val((Mid(CadenaLimpiA, I, 1))) * Z)
            Z = Z + 1
        End If
        I = I - 1
    Wend
    
    Z = 11 - (Val(CadenA) - Int((Val(CadenA)) / 11) * 11)
    
    XXXX = Asc(DiG)
        If DiG = 0 And Z = 11 Then
            EsRUT = True
        Else
                If Z = DiG Then
                    EsRUT = True
                Else
                    EsRUT = False
                End If
        End If
Else
    EsRUT = False
End If
CadenA = Empty
CadenaLimpiA = Empty
End Function

