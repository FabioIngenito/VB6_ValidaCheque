VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AbreBanco()
   Set adoCMC7 = New ADODB.Connection
   adoCMC7.CursorLocation = adUseClient
   'adoCMC7.ConnectionString = "PROVIDER=Microsoft.Jet.Oledb.4.0; Persist Security Info=false; Data Source= " & strPath & "bdbancos.mdb"
   adoCMC7.ConnectionString = "Provider=Microsoft.Ace.Oledb.12.0; Persist Security Info=False; Data Source= " & strPath & "bdbancos.mdb"
   adoCMC7.Open
End Sub

Public Sub FechaBanco()
   adoCMC7.Close
   Set adoCMC7 = Nothing
End Sub

Sub CommandText(ByVal texto As String)
'rotina que seta os parâmetros do objeto command. Ela recebe o texto do command e o executa
    Dim cmdtexto As Command
    Set cmdtexto = New Command
    cmdtexto.ActiveConnection = adoCMC7
    cmdtexto.CommandType = adCmdText
    cmdtexto.CommandText = texto
    cmdtexto.Execute
    Set cmdtexto = Nothing
End Sub

Function CommandTextRetorna(ByVal texto As String) As ADODB.Recordset
'rotina que seta os parâmetros do objeto command. Ela recebe o texto do command e o executa
    Dim cmdtexto As ADODB.Command
    Set cmdtexto = New ADODB.Command
    cmdtexto.ActiveConnection = adoCMC7
    cmdtexto.CommandType = adCmdText
    cmdtexto.CommandText = texto
    Set CommandTextRetorna = cmdtexto.Execute
    Set cmdtexto = Nothing
End Function

Function RetornaExistencia(ByVal SQL As String) As Boolean
'rotina que verifica se já existe a Chave Primária ou
'qq dado do banco dependendo da string passada
Dim rstTexto As ADODB.Recordset

    Set rstTexto = New ADODB.Recordset
    rstTexto.Open SQL, adoCMC7, adOpenForwardOnly, adLockReadOnly

    If rstTexto.BOF And rstTexto.EOF Then
        RetornaExistencia = False
    Else
        RetornaExistencia = True
    End If

    rstTexto.Close
    Set rstTexto = Nothing
End Function

Function ContaExistencia(ByVal SQL As String) As Long
'rotina que conta as Chave Primária ou
'qq dado do banco dependendo da string passada
Dim rstTexto As ADODB.Recordset

    Set rstTexto = New ADODB.Recordset
    rstTexto.Open SQL, adoCMC7, adOpenForwardOnly, adLockReadOnly

    If rstTexto.BOF And rstTexto.EOF Then
        ContaExistencia = 0
    Else
        ContaExistencia = rstTexto.RecordCount
    End If

    rstTexto.Close
    Set rstTexto = Nothing
End Function

Function RetornaTexto(ByVal SQL As String, _
                      ByVal SoCodigo As Boolean) As String
'rotina que verifica se já existe a Chave Primária ou
'qq dado do banco dependendo da string passada
Dim rstTexto As ADODB.Recordset

   Set rstTexto = New ADODB.Recordset
   rstTexto.Open SQL, adoCMC7, adOpenForwardOnly, adLockReadOnly

   If rstTexto.BOF And rstTexto.EOF Then
      RetornaTexto = ""
   Else
   
      If SoCodigo Then
         RetornaTexto = rstTexto.Fields(0).Value
      Else
         RetornaTexto = rstTexto.Fields(1).Value & " - " & _
                        rstTexto.Fields(0).Value
      End If

   End If

   rstTexto.Close
   Set rstTexto = Nothing
End Function

Function PreencheSQL(ByVal strBD As String) As String

    PreencheSQL = "SELECT BD.c_codigo, BD.c_descricao " & _
                  "FROM " & strBD & " BD " & _
                  "ORDER BY BD.c_descricao"

End Function
