VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Contador As Integer 'Não pode ser tipo BYTE!

Public Function Modulo10(Calculo As Variant)
Dim i As Integer
Dim nDigito As Integer
Dim nSoma As Integer
Dim nSoma2 As Integer

nSoma = 0
nDigito = 0
Calculo = Trim(Calculo)

If Len(Calculo) <> 0 Then
    
    For i = 0 To Len(Calculo) - 1
        nSoma2 = (Val(Mid(Calculo, Len(Calculo) - i, 1)) * 2)
        If Len(nSoma2) = 2 Then
            nSoma2 = Val(Mid(nSoma2, 1, 1)) + Val(Mid(nSoma2, 2, 1))
        End If
        
        nSoma = nSoma + nSoma2
        
        i = i + 1
        If i = Len(Calculo) Then
            Exit For
        End If
        
        nSoma = nSoma + (Val(Mid(Calculo, Len(Calculo) - i, 1)))
    Next
    
    nDigito = nSoma Mod 10
    If nDigito <> 0 Then
       nDigito = 10 - nDigito
    End If
    
End If

Modulo10 = nDigito

End Function
Public Function Modulo11(Calculo As Variant)
Dim i As Integer
Dim a As Integer
Dim nDigito As Integer
Dim nSoma As Integer
Dim nSoma2 As Integer

nSoma = 0
nDigito = 0
Calculo = Trim(Calculo)

If Len(Calculo) <> 0 Then
    
    While i <> Len(Calculo)
    
        For a = 2 To 9
            
                       
            nSoma2 = (Val(Mid(Calculo, Len(Calculo) - i, 1)) * a)
        
            'If Len(nSoma2) = 2 Then
            '    nSoma2 = Val(Mid(nSoma2, 1, 1)) + Val(Mid(nSoma2, 2, 1))
            'End If
        
            nSoma = nSoma + nSoma2
           
            i = i + 1
            
            If i = Len(Calculo) Then Exit For
                                       
        Next a
        
    Wend
    
    nDigito = nSoma Mod 11
    If nDigito <> 0 Then
    
       nDigito = 11 - nDigito
       'If nDigito = 10 Then
       '     nDigito = 0
       'End If
       If nDigito = 0 Or nDigito = 1 Or nDigito > 9 Then
          nDigito = 1
       End If
       
'mf em 27022002 10:50hs
Else
    nDigito = 1
       
    End If
    
End If

Modulo11 = nDigito

End Function

Public Function Modulo11Arr(Calculo As Variant)
Dim i As Integer
Dim a As Integer
Dim nDigito As Integer
Dim nSoma As Integer
Dim nSoma2 As Integer

nSoma = 0
nDigito = 0
Calculo = Trim(Calculo)

If Len(Calculo) <> 0 Then
    
    While i <> Len(Calculo)
    
        For a = 2 To 9
            
                       
            nSoma2 = (Val(Mid(Calculo, Len(Calculo) - i, 1)) * a)
        
           ' If Len(nSoma2) = 2 Then
           '     nSoma2 = Val(Mid(nSoma2, 1, 1)) + Val(Mid(nSoma2, 2, 1))
           ' End If
        
            nSoma = nSoma + nSoma2
           
            i = i + 1
            
            If i = Len(Calculo) Then Exit For
                                       
        Next a
        
    Wend

''' Conforme consta no manual Febraban versão 4  arrecadação datado em 28/05/2004
''' o dv geral deverá seguir a seguinte regra
    
''A soma dos produtos dessa multiplicação é dividida por 11 e
''obtém-se o resto da divisão, este resto deve ser subtraído de 11, o produto da subtração é o DAC - 'Dígito'.
''Observação: Quando o resto da divisão for igual a 0 ou 1, atribuí-se ao DV o digito "0",
''e quando for 10, atribuí-se ao DV o digito "1".
    

        nDigito = nSoma Mod 11
        nDigito = 11 - nDigito

        If nDigito > 9 Then
            nDigito = 0
        End If
      
        If nDigito = 10 Then
            nDigito = 1
        End If
    
End If


Modulo11Arr = nDigito

End Function

Public Function MOD10(strDACC As String) As String
Dim Marcador As Boolean
Dim intDACC As Integer
Dim intDACCTotal As Integer
Dim intDACCResto As Integer

    intDACC = 0
    intDACCTotal = 0
    Marcador = True
    
    For Contador = Len(strDACC) To 1 Step -1
        intDACC = Val(Mid(strDACC, Contador, 1))
    
        If Marcador = True Then
            Marcador = False
            intDACC = intDACC * 2
    
            If intDACC > 9 Then
                intDACC = Val(Mid(Trim(Str(intDACC)), 1, 1)) + Val(Mid(Trim(Str(intDACC)), 2, 1))
            End If
    
        Else
            Marcador = True
        End If
    
        intDACCTotal = intDACCTotal + intDACC
    Next
    
    intDACCResto = Int(intDACCTotal Mod 10)
    
    If intDACCResto = 0 Then
        MOD10 = 0
    Else
        MOD10 = 10 - intDACCResto
    End If

End Function

Public Function MOD11(strDACB As String) As String
Dim intDACB As Integer
Dim intDACBTotal As Integer
Dim intDACBResto As Integer
Dim Peso As Byte

    intDACB = 0
    intDACBTotal = 0
    Peso = 2
    
    For Contador = Len(strDACB) To 1 Step -1
        intDACB = Val(Mid(strDACB, Contador, 1))
        intDACB = intDACB * Peso
        intDACBTotal = intDACBTotal + intDACB
        Peso = Peso + 1
        
        If Peso > 9 Then
            Peso = 2
        End If
    
    Next
    
    intDACBResto = Int(intDACBTotal Mod 11)
    
    If intDACBResto = 0 Or intDACBResto = 1 Or intDACBResto = 10 Or intDACBResto = 11 Then
        MOD11 = 1
    Else
        MOD11 = 11 - intDACBResto
    End If

End Function

Public Function DAC(strNumeroDaAgenciaCedente As String, strNumeroDaContaCorrente As String, strCarteira As String, strNossoNumero As String) As String
Dim Marcador As Boolean
Dim strDACD As String
Dim intDACD As Integer
Dim intDACDTotal As Integer
Dim intDACDResto As Integer

    intDACD = 0
    intDACDTotal = 0
    Marcador = True

    If strCarteira = "126" Or strCarteira = "131" Or _
       strCarteira = "146" Or strCarteira = "150" Or _
       strCarteira = "168" Then
        strDACD = strCarteira & strNossoNumero
    Else
        strDACD = strNumeroDaAgenciaCedente & strNumeroDaContaCorrente & _
                  strCarteira & strNossoNumero
    End If

    For Contador = Len(strDACD) To 1 Step -1
        intDACD = Val(Mid(strDACD, Contador, 1))
    
        If Marcador = True Then
            Marcador = False
            intDACD = intDACD * 2
    
            If intDACD > 9 Then
                intDACD = Val(Mid(Trim(Str(intDACD)), 1, 1)) + Val(Mid(Trim(Str(intDACD)), 2, 1))
            End If
    
        Else
            Marcador = True
        End If
    
        intDACDTotal = intDACDTotal + intDACD
    Next
    
    intDACDResto = Int(intDACDTotal Mod 10)
    
    If intDACDResto = 0 Then
        DAC = 0
    Else
        DAC = 10 - intDACDResto
    End If

End Function

Public Function DV_MOD10(Numero As String)
    Dim tamanho As Integer
    Dim Numerox As String
    Dim soma As Long
    Dim Somax As Long
    Dim Peso As String
    Dim i As Integer
    
    tamanho = Len(Numero)
    Numerox = Numero
    soma = 0
    
    If tamanho <> 0 Then
        Peso = Right("1212121212", tamanho)
        
        For i = tamanho To 1 Step -1
            Somax = Val(Mid(Numerox, i, 1)) * Val(Mid(Peso, i, 1))
            If Somax > 9 Then
                Somax = Val(Mid(Somax, 1, 1)) + Val(Mid(Somax, 2, 1))
            End If
            soma = soma + Somax
            If i > 1 Then
               Numerox = Left(Numerox, (i - 1))
            End If
        Next
        
        soma = 9 * soma
        DV_MOD10 = soma Mod (10)
    End If

End Function

Public Function Dig_Base10(xPCODIGO)
   
  Dim cMaxDig As String
  Dim flag As Boolean
  Dim Tx As Integer
  Dim TVlr As String
  Dim TDig As String
   
  cMaxDig = Len(xPCODIGO)
  TVlr = 0
  flag = True
  For Tx = cMaxDig To 1 Step -1
    If flag Then
      TDig = (Val(Mid(xPCODIGO, Tx, 1)) * 2)
      flag = False
    Else
      TDig = Val(Mid(xPCODIGO, Tx, 1))
      flag = True
    End If
       
    If TDig > 9 Then
      TDig = 1 + (TDig - 10)
      TVlr = Val(TVlr) + Val(TDig)
    Else
      TVlr = Val(TVlr) + Val(TDig)
    End If
  Next
   
  TDig = 10 * ((TVlr / 10) - Int(TVlr / 10))

  If TDig > 0 Then
    TDig = 10 - TDig
  End If

  Dig_Base10 = TDig

End Function
