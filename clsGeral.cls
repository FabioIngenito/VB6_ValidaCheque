VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub LimpaCampos(ByVal frm As Form)
'Coloca um espaço em branco em todas as caixas de texto do formulário
Dim obj As Object
 
   For Each obj In frm.Controls
        
      If (TypeOf obj Is TextBox) Or (TypeOf obj Is ComboBox) Then 'Or (TypeOf obj Is MaskEdBox) Then
         obj.Text = ""
      'ElseIf TypeOf obj Is DTPicker Then
      '   obj.Value = Now
      End If

   Next

End Sub

Public Sub ConfiguraBotoes(ByVal frm As Form, _
                           trava As Boolean)
'Configura Enable em todas os botões do formulário
Dim obj As Object
 
   For Each obj In frm.Controls
        
      If (TypeOf obj Is CommandButton) Then
         obj.Enabled = trava
      End If

   Next

End Sub

Public Sub ConfiguraCampos(ByVal frm As Form, _
                           trava As Boolean, _
                           cor As OLE_COLOR)
'Configura Locked e ForeColor em todas as caixas de texto do formulário
Dim obj As Object
 
   For Each obj In frm.Controls
        
      If (TypeOf obj Is TextBox) Or (TypeOf obj Is ComboBox) Then
         obj.ForeColor = cor
         obj.Locked = trava
      'ElseIf TypeOf obj Is MaskEdBox Then
      '   obj.ForeColor = cor
      '   obj.Enabled = Not trava
      'ElseIf TypeOf obj Is DTPicker Then
      '   obj.CalendarForeColor = cor
      '   obj.CalendarTitleForeColor = cor
      End If

   Next

End Sub

Public Function TrataApostrofe(ByVal strTexto As String) As String
Dim intPosicao As Integer
Dim intLocal As Integer

    intPosicao = 1
    intLocal = InStr(intPosicao, strTexto, "'")
    strTexto = Trim(strTexto)

    Do While intPosicao <= intLocal
        TrataApostrofe = TrataApostrofe + Mid(strTexto, intPosicao, intLocal - intPosicao + 1) + "'"
        intPosicao = intLocal + 1
        intLocal = InStr(intPosicao, strTexto, "'")
    Loop

    TrataApostrofe = TrataApostrofe + Mid(strTexto, intPosicao)
End Function

Public Function F_Centraliza(ByVal Formulario As Form)
    Formulario.Left = (Screen.Width - Formulario.Width) / 2
    Formulario.Top = (Screen.Height - Formulario.Height) / 2 - 1080
End Function

Public Function AcertaCodigo(ByVal strCodigo As String, _
                             ByVal intTamanho As Integer, _
                             ByVal strCaracter As String) As String
    strCodigo = Right(strCodigo, intTamanho)
    AcertaCodigo = String(intTamanho - Len(strCodigo), strCaracter) & strCodigo
End Function

Public Function PoeAspa(strString As String, strTipo As String) As String

    If strTipo = "A" Then
        PoeAspa = "'" & strString & "'"
    Else
        PoeAspa = strString
    End If

End Function

Public Function PegaAposTraco(strTexto As String) As String
    
    If Not strTexto = "" Then
        PegaAposTraco = Right(strTexto, Len(strTexto) - InStrRev(strTexto, "-") - 1)
    End If
    
End Function

Public Function PegaAposTraco2(strTexto As String) As String
Dim arrPalavra() As String
' O mesmo retorno que o PegaAposTraço mas de uma outra maneira ... by Neusa

    If Not strTexto = "" Then
        arrPalavra = Split(strTexto, "-")
        PegaAposTraco2 = Trim(arrPalavra(1))
    End If
    
End Function

'Função para colocar o ENTER como TAB
Public Sub TabXEnter(Key As Integer)
    If Key = 13 Then Key = 0: SendKeys "{TAB}"
End Sub

'Função que ajuda a preencher uma ComboBox
Function AutoSel(Cmb As ComboBox, KeyCode As Integer)
Dim Text As String
Dim i As Long
Dim Temp As String
   
    'Debug.Print KeyCode
    
    If KeyCode = vbEnter Then Exit Function
    If KeyCode = 8 Then Exit Function   'Backspace
    If KeyCode = 37 Then Exit Function  'left key
    If KeyCode = 38 Then Exit Function  'up arrow key
    If KeyCode = 39 Then Exit Function  'right key
    If KeyCode = 40 Then Exit Function  'down arrow key
    If KeyCode = 46 Then Exit Function  'delete key
    If KeyCode = 33 Then Exit Function  'page up key
    If KeyCode = 34 Then Exit Function  'page down key
    If KeyCode = 35 Then Exit Function  'end key
    If KeyCode = 36 Then Exit Function  'home key

    Text = Cmb.Text
    
    For i = 0 To Cmb.ListCount - 1
        Temp = Left(Cmb.List(i), Len(Text))
        
        If LCase(Temp) = LCase(Text) Then
            Cmb.Text = Cmb.List(i)
            Cmb.ListIndex = i
            Cmb.SelStart = Len(Text)
            Cmb.SelLength = Len(Cmb.List(i))
            'Cmb.SetFocus
        End If

    Next

End Function

'http://www.macoratti.net/d250902.htm
'Copiar para área de transferência
Public Sub AreaTransferencia(txt As String)
    Clipboard.Clear ' limpa o ClibBoard
    Clipboard.SetText txt ' Poe o texto no ClipBoard
End Sub
