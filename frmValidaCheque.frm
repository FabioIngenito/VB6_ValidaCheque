VERSION 5.00
Begin VB.Form frmValidaCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valida Cheque - CMC7 - (Caracteres Magnéticos Codificados em 7 Barras)"
   ClientHeight    =   8505
   ClientLeft      =   1155
   ClientTop       =   2130
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdAreaTransferenciaTerceiraBanda 
      Caption         =   "Área Transferência"
      Height          =   495
      Left            =   180
      TabIndex        =   37
      Top             =   6540
      Width           =   1275
   End
   Begin VB.CommandButton cmdAreaTransferenciaSegundaBanda 
      Caption         =   "Área Transferência"
      Height          =   495
      Left            =   180
      TabIndex        =   36
      Top             =   5160
      Width           =   1275
   End
   Begin VB.CommandButton cmdAreaTransferenciaPrimeiraBanda 
      Caption         =   "Área Transferência"
      Height          =   495
      Left            =   180
      TabIndex        =   35
      Top             =   3780
      Width           =   1275
   End
   Begin VB.CommandButton cmdAreaTransferenciaSemCaracteresEspeciais 
      Caption         =   "A.Transf."
      Height          =   315
      Left            =   7140
      TabIndex        =   34
      Top             =   7560
      Width           =   795
   End
   Begin VB.CommandButton cmdAreaTransferenciaComCaracteresEspeciais 
      Caption         =   "A.Transf."
      Height          =   315
      Left            =   7140
      TabIndex        =   33
      Top             =   7200
      Width           =   795
   End
   Begin VB.TextBox txtBandaMagneticaSemCaracteresEspeciais 
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   34
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "001000580012189175704001686452"
      Top             =   7560
      Width           =   3135
   End
   Begin VB.ComboBox cboBanco 
      ForeColor       =   &H00008000&
      Height          =   315
      ItemData        =   "frmValidaCheque.frx":0000
      Left            =   1740
      List            =   "frmValidaCheque.frx":0002
      TabIndex        =   2
      Text            =   "237 - BANCO BRADESCO S.A."
      Top             =   3060
      Width           =   3015
   End
   Begin VB.CommandButton cmdLimparCampos 
      Caption         =   "&Limpar Campos"
      Height          =   315
      Left            =   5280
      TabIndex        =   29
      Top             =   8100
      Width           =   1275
   End
   Begin VB.CommandButton cmdExemplo3 
      Caption         =   "Exemplo &3"
      Height          =   315
      Left            =   2700
      TabIndex        =   28
      Top             =   8100
      Width           =   1275
   End
   Begin VB.TextBox txtBandaMagneticaComCaracteresEspeciais 
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   34
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "<00100058<0012189175>704001686452>"
      Top             =   7200
      Width           =   3135
   End
   Begin VB.CommandButton cmdExemplo2 
      Caption         =   "Exemplo &2"
      Height          =   315
      Left            =   1380
      TabIndex        =   27
      Top             =   8100
      Width           =   1275
   End
   Begin VB.CommandButton cmdExemplo1 
      Caption         =   "Exemplo &1"
      Height          =   315
      Left            =   60
      TabIndex        =   26
      Top             =   8100
      Width           =   1275
   End
   Begin VB.ComboBox cboTipificacaoCheque 
      ForeColor       =   &H00008000&
      Height          =   315
      ItemData        =   "frmValidaCheque.frx":0004
      Left            =   1740
      List            =   "frmValidaCheque.frx":0017
      TabIndex        =   14
      Text            =   "5 - Comum"
      Top             =   5280
      Width           =   1635
   End
   Begin VB.TextBox txtBandaDg3 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   6660
      Width           =   255
   End
   Begin VB.TextBox txtBandaDg2 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "2"
      Top             =   3900
      Width           =   255
   End
   Begin VB.TextBox txtBandaDg1 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   5820
      Width           =   255
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      Height          =   315
      Left            =   6600
      TabIndex        =   30
      Top             =   8100
      Width           =   1275
   End
   Begin VB.TextBox txtFaixa3 
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "2000001044649"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtFaixa2 
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0180032515"
      Top             =   4740
      Width           =   1035
   End
   Begin VB.TextBox txtFaixa1 
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "00125402"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtChequeNumero 
      Height          =   315
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "003489"
      ToolTipText     =   "Número do Cheque"
      Top             =   4860
      Width           =   675
   End
   Begin VB.TextBox txtConta 
      Height          =   315
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   20
      Text            =   "7750436240"
      ToolTipText     =   "Número da Conta Corrente"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtAgencia 
      Height          =   315
      Left            =   1740
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "3184"
      ToolTipText     =   "Código da Agência"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txtComp 
      Height          =   315
      Left            =   1740
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "001"
      ToolTipText     =   "Código da Câmara de Compensação"
      Top             =   4440
      Width           =   435
   End
   Begin VB.Label lblBandaMagneticaSemCarcteresEspeciais 
      BackColor       =   &H00C00000&
      Caption         =   "Banda Magnética (30 sem os 4 caracteres especiais)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   31
      Top             =   7560
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   480
      Picture         =   "frmValidaCheque.frx":0060
      Top             =   120
      Width           =   7005
   End
   Begin VB.Line Line6 
      X1              =   60
      X2              =   7860
      Y1              =   7980
      Y2              =   7980
   End
   Begin VB.Label lblBandaMagneticaComCarcteresEspeciais 
      BackColor       =   &H00C00000&
      Caption         =   "Banda Magnética (34 com os 4 caracteres especiais)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   24
      Top             =   7260
      Width           =   3795
   End
   Begin VB.Line Line5 
      X1              =   1620
      X2              =   1620
      Y1              =   7080
      Y2              =   3060
   End
   Begin VB.Line Line3 
      X1              =   60
      X2              =   7860
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line4 
      X1              =   4860
      X2              =   4860
      Y1              =   7080
      Y2              =   3060
   End
   Begin VB.Line Line2 
      X1              =   180
      X2              =   7860
      Y1              =   5700
      Y2              =   5700
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   7860
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblTerceiraBanda 
      BackColor       =   &H000000C0&
      Caption         =   "Terceira Banda (12)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   5820
      Width           =   1515
   End
   Begin VB.Label lblSegundaBanda 
      BackColor       =   &H000000C0&
      Caption         =   "Segunda Banda (10)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label lblPrimeiraBanda 
      BackColor       =   &H000000C0&
      Caption         =   "Primeira Banda (8)"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   3060
      Width           =   1515
   End
   Begin VB.Label lblTipificacaoCheque 
      Caption         =   "Tipificação do Cheque (1)"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4980
      TabIndex        =   15
      ToolTipText     =   "Tipificação do Cheque (5–Comum 6–Bancário 7–Salário 8–Administrativo 9–CPMF)"
      Top             =   5340
      Width           =   1935
   End
   Begin VB.Label lblDV3 
      Caption         =   "Dígito verificador da Terceira Banda (1)"
      Height          =   255
      Left            =   4980
      TabIndex        =   23
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label lblDV2 
      Caption         =   "Dígito Verificador da Segunda Banda (1)"
      Height          =   255
      Left            =   4980
      TabIndex        =   7
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label lblDV1 
      Caption         =   "Dígito verificador da Primeira Banda (1)"
      Height          =   255
      Left            =   4980
      TabIndex        =   19
      Top             =   5880
      Width           =   2835
   End
   Begin VB.Label lblChequeNumero 
      Caption         =   "Número do Cheque (6)"
      Height          =   255
      Left            =   4980
      TabIndex        =   13
      ToolTipText     =   "Número do Cheque"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblConta 
      Caption         =   "Conta (10)"
      Height          =   255
      Left            =   4980
      TabIndex        =   21
      ToolTipText     =   "Número da Conta Corrente"
      Top             =   6300
      Width           =   795
   End
   Begin VB.Label lblAgencia 
      Caption         =   "Código da Agência (4)"
      Height          =   255
      Left            =   4980
      TabIndex        =   5
      ToolTipText     =   "Código da Agência"
      Top             =   3540
      Width           =   1755
   End
   Begin VB.Label lblBanco 
      Caption         =   "Código do Banco (3)"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4980
      TabIndex        =   3
      ToolTipText     =   "Código do Banco"
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label lblComp 
      Caption         =   "Código da Câmara de Compensação (3)"
      Height          =   255
      Left            =   4980
      TabIndex        =   11
      ToolTipText     =   "Código da Câmara de Compensação"
      Top             =   4500
      Width           =   2835
   End
End
Attribute VB_Name = "frmValidaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tarja As String
Dim ChecaCMC7 As Boolean
Dim MsgErroLog As String

Private Sub cboBanco_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub cboBanco_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(cboBanco.Text) Then Geral.AutoSel cboBanco, KeyCode
End Sub

Private Sub cmdAreaTransferenciaComCaracteresEspeciais_Click()
    Geral.AreaTransferencia txtBandaMagneticaComCaracteresEspeciais.Text
End Sub

Private Sub cmdAreaTransferenciaSemCaracteresEspeciais_Click()
    Geral.AreaTransferencia txtBandaMagneticaSemCaracteresEspeciais.Text
End Sub

Private Sub cmdAreaTransferenciaPrimeiraBanda_Click()
    Geral.AreaTransferencia txtFaixa1.Text
End Sub

Private Sub cmdAreaTransferenciaSegundaBanda_Click()
    Geral.AreaTransferencia txtFaixa2.Text
End Sub

Private Sub cmdAreaTransferenciaTerceiraBanda_Click()
    Geral.AreaTransferencia txtFaixa3.Text
End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtChequeNumero_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtComp_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub cboTipificacaoCheque_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub cboTipificacaoCheque_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(cboTipificacaoCheque.Text) Then Geral.AutoSel cboTipificacaoCheque, KeyCode
End Sub

Private Sub txtConta_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo TrataErro
'Dimensiona as RecordSets
Dim rsBD As ADODB.Recordset
Dim rsTCheque As ADODB.Recordset

    'Limpa a ComboBox
    cboBanco.Clear

    'Instancia "Class Module" clsBD
    Set BD = New clsBD

    'Instacia a Conexão ao Banco de Dados
    Set adoCMC7 = New ADODB.Connection
    
    'Abrir banco
    BD.AbreBanco

    'Instancia a RecordSet rsBD
    Set rsBD = New ADODB.Recordset
    strSQL = BD.PreencheSQL("tbl_bancos")
    Set rsBD = BD.CommandTextRetorna(strSQL)

    'Carregar Base de Dados de "Bancos" na ComboBox
    Do While rsBD.EOF = False
        cboBanco.AddItem rsBD.Fields("c_codigo").Value & " - " & _
                         rsBD.Fields("c_descricao").Value
        cboBanco.ItemData(cboBanco.ListCount - 1) = rsBD.Fields("c_codigo").Value
        rsBD.MoveNext
    Loop

    'Destruir a instância
    Set rsBD = Nothing
    
    'Limpa a ComboBox
    cboTipificacaoCheque.Clear
    
    'Instancia a RecordSet rsBD
    Set rsTCheque = New ADODB.Recordset
    strSQL = BD.PreencheSQL("tbl_tipificacaoCheque")
    Set rsTCheque = BD.CommandTextRetorna(strSQL)
    
    'Carregar Base de Dados de "Tipificação de Cheque" na ComboBox
    Do While rsTCheque.EOF = False
        cboTipificacaoCheque.AddItem rsTCheque.Fields("c_codigo").Value & " - " & _
                                     rsTCheque.Fields("c_descricao").Value
        cboTipificacaoCheque.ItemData(cboTipificacaoCheque.ListCount - 1) = rsTCheque.Fields("c_codigo").Value
        rsTCheque.MoveNext
    Loop
    
    BD.FechaBanco
    
    cmdLimparCampos_Click
       
    Exit Sub
       
TrataErro:
    MsgBox "Ocorreu um erro!", vbCritical + vbOKOnly, "ERRO"

End Sub

Private Sub cmdLimparCampos_Click()
On Error GoTo TrataErro

    'Primeira Banda
    txtFaixa1.Text = ""
    cboBanco.ListIndex = -1
    cboBanco.Text = ""
    txtAgencia.Text = ""
    txtBandaDg2.Text = ""
    
    'Segunda Banda:
    txtFaixa2.Text = ""
    txtComp.Text = ""
    txtChequeNumero.Text = ""
    cboTipificacaoCheque.ListIndex = 2  '5 - Comum

    'Terceira Banda:
    txtFaixa3.Text = ""
    txtBandaDg1.Text = ""
    txtConta.Text = ""
    txtBandaDg3.Text = ""
    
    'Banda Magnética:
    txtBandaMagneticaComCaracteresEspeciais.Text = ""
    txtBandaMagneticaSemCaracteresEspeciais.Text = ""
    Exit Sub

TrataErro:
    MsgBox "Ocorreu um erro!", vbCritical + vbOKOnly, "ERRO"

End Sub

Private Sub cmdCalcular_Click()

    'Validar a combo "cboBanco"
    If cboBanco.ListIndex = -1 Then
        MsgBox "Por favor, escolha um Banco.", vbCritical + vbOKOnly, "BANCO"
        cboBanco.Text = ""
        cboBanco.SetFocus
        Exit Sub
    End If

    'Validar a combo "cboTipificacaoCheque"
    If cboTipificacaoCheque.ListIndex = -1 Then
        MsgBox "Por favor, escolha um Tipo de Cheque.", vbCritical + vbOKOnly, "TIPIFICAÇÃO DO CHEQUE"
        cboTipificacaoCheque.Text = ""
        cboTipificacaoCheque.SetFocus
        Exit Sub
    End If
    
    'Instancia a "Class Module" clsModulo
    Set modul = New clsModulo

    'Acerta se cada campo está no tamanho correto
    AcertaTamanhoCampo

    txtFaixa1.Text = Format(cboBanco.ItemData(cboBanco.ListIndex), "000") & txtAgencia.Text
    txtFaixa2.Text = txtComp.Text & txtChequeNumero.Text & cboTipificacaoCheque.ItemData(cboTipificacaoCheque.ListIndex)
    txtFaixa3.Text = txtConta.Text

    '1 - Calcular PRIMEIRO o dígito Verificar do Segunda Banda!
    'txtBandaDg2.Text = modul.Modulo10(txtFaixa2.Text)
    'txtBandaDg2.Text = modul.MOD10(txtFaixa2.Text)
    'txtBandaDg2.Text = modul.DV_MOD10(txtFaixa2.Text)
    txtBandaDg2.Text = modul.Dig_Base10(txtFaixa2.Text)

    '2 - Fazer a Primeira Banda
    'txtBandaDg1.Text = modul.Modulo10(txtFaixa1.Text)
    'txtBandaDg1.Text = modul.MOD10(txtFaixa1.Text)
    'txtBandaDg1.Text = modul.DV_MOD10(txtFaixa1.Text)
    txtBandaDg1.Text = modul.Dig_Base10(txtFaixa1.Text)
      
    '3 - Fazer a terceira Banda
    'txtBandaDg3.Text = modul.DV_MOD10(Format(txtFaixa3.Text, "0000000000"))
    txtBandaDg3.Text = modul.DV_MOD10(txtFaixa3.Text)
   
    'Colocando os dígitos nas faixas (Somente faixas 1 e 3):
    txtFaixa1.Text = txtFaixa1.Text & txtBandaDg2.Text
    txtFaixa1.Text = Format(txtFaixa1.Text, "00000000")
    
    txtFaixa3.Text = txtBandaDg1.Text & txtFaixa3.Text & txtBandaDg3.Text
    
    MontaBandaMagnetica
End Sub

Private Sub AcertaTamanhoCampo()
    txtAgencia.Text = Geral.AcertaCodigo(txtAgencia.Text, 4, "0")
    txtComp.Text = Geral.AcertaCodigo(txtComp.Text, 3, "0")
    txtChequeNumero.Text = Geral.AcertaCodigo(txtChequeNumero.Text, 6, "0")
    txtConta.Text = Geral.AcertaCodigo(txtConta.Text, 10, "0")
End Sub

Private Sub MontaBandaMagnetica()
'======================================================
'http://www.devmedia.com.br/forum/layout-cheque/285633
'======================================================
'Pessoal dei uma lida num artigo do Aroldo Zanela sobre layout da banda magnética, mas to com muitas dúvidas. Por exemplo eu tenho um cheque com os seguintes dados( Banco do Brasil)
'
'Comp:001
'Banco:001
'Agencia:0005
'DV:1
'C1:8
'Conta:0016864-5
'C2:3
'Série:001
'Cheque Nº:218917
'C3:8
'
'E a banda magnética da seguinte maneira:
'<00100058<0012189175<704001686452>
'
'Não batem os campos, alguem se habilita?
'
'Daniel Miranda Cruz

' ----- ATENÇÃO!!! ------------------ EXPLICAÇÃO!!!

'E a banda magnética da seguinte maneira:
'<00100058<0012189175<704001686452>

'<001 = banco
'0005 = agência
'8 = dv2
'<001 = câmara de compensação (comp)
'218917 = número do cheque
'5 = tipificação ( 5 normal, 6 bancário, 7 salário, 8 administrativo, etc)
'>7 = dv1
'0400168645 = número da conta (*)
'2> = dv3

'*No número da conta do BB é colocado 2 dígitos para identificar o tipo de conta, neste caso 04. Esta é única diferença entre a banda magnética e a linha superior.

'Outros bancos:
'Itaú, eles colocam 4 dígitos, que é diferente para cada cheque do talão.
'Unibanco faz o mesmo, porém são 3 dígitos.
'Bradesco: número fixo 775.

'A conta da banda magnética possui tem 10 caracteres, atualmente os bancos estão utilizando parte da conta para colocar mais segurança, como o Itaú que criou mais 4 dvs e colocou nas 4 primeiras posições do número da conta.

'======================================================
'mais detalhes,
'http://www.veloso.adm.br
'======================================================

    txtBandaMagneticaComCaracteresEspeciais.Text = "<" & _
                             Format(cboBanco.ItemData(cboBanco.ListIndex), "000") & _
                             txtAgencia.Text & _
                             txtBandaDg2.Text & _
                             "<" & _
                             txtComp.Text & _
                             txtChequeNumero.Text & _
                             cboTipificacaoCheque.ItemData(cboTipificacaoCheque.ListIndex) & _
                             ">" & _
                             txtBandaDg1.Text & _
                             txtConta.Text & _
                             txtBandaDg3.Text & _
                             ">"

    txtBandaMagneticaSemCaracteresEspeciais.Text = Format(cboBanco.ItemData(cboBanco.ListIndex), "000") & _
                             txtAgencia.Text & _
                             txtBandaDg2.Text & _
                             txtComp.Text & _
                             txtChequeNumero.Text & _
                             cboTipificacaoCheque.ItemData(cboTipificacaoCheque.ListIndex) & _
                             txtBandaDg1.Text & _
                             txtConta.Text & _
                             txtBandaDg3.Text

'------------------------------------------------------------------------------
    'CMC7 CERTO
    '001254090180032515170010446428
'------------------------------------------------------------------------------
End Sub
    
Private Sub cmdExemplo1_Click()
'======================================================
'Exemplo que funciona:
'http://www.veloso.adm.br/leiautecmc7.html
'======================================================
'-------------------------------------------
'CMC7 - exemplo:
'<23731842<0010034895>777504362409>
'-------------------------------------------
'a - 237        - Número do Banco
'b - 3184       - Número da Agência
'c - 2          - Dígito verificador da Comp+Cheque+Tipificação
'd - 001        - Comp (câmara de compensação - 018-SP, 001-RJ, etc)
'e - 003489     - Número do cheque
'f - 5          - Tipificação(5-Comum 6-Bancário 7-Salário 8-Administr. 9-CPMF)
'g - 7          - Dígito verificador do Banco+Agência
'h - 7750436240 - Número da Conta
'i - 9          - Dígito verificador da Conta*
'* O Dv da Conta no CMC7 não tem nada haver com o Dv da Conta na Linha1, este inclusive faz parte do número da conta no CMC7.
'-------------------------------------------

    cmdLimparCampos_Click
    cboBanco.ListIndex = 20             'a '237 - BANCO BRADESCO S.A.
    txtAgencia.Text = "3184"            'b
    txtComp.Text = "001"                'd
    txtChequeNumero.Text = "003489"     'e
    cboTipificacaoCheque.ListIndex = 2  'f '5 - Comum
    txtConta.Text = "7750436240"        'h
    cmdCalcular_Click
End Sub

Private Sub cmdExemplo2_Click()
'-------------------------------------------
'Exemplo que funciona:
'http://www.fesppr.br/~erico/x%202009%20ASIG/2009%20-%20sala%20305/Trabalho%20CMC7_%20Cristiane%20Faria,%20Fernanda,%20Ana%20Paula,%20Francine%20e%20Tatiane.pdf
'-------------------------------------------
'Os separadores dividem o CMC7 em três bandas. Exemplo:
'Primeira Banda     Segunda Banda   Terceira Banda
'23704948           0180017935      377506100112
'-------------------------------------------
'Primeira Banda - 237 0494 8
'237 Código do Banco
'0494 Código da Agência
'8 Dígito verificador da Segunda Banda
'-------------------------------------------
'Segunda Banda - 018 001793 5
'018 Código da câmara de Compensação
'001793 Número do cheque
'5 - Tipificação do cheque (5 – Comum 6-Bancário 7 – Salário 8–Administrativo 9 – CPMF)
'-------------------------------------------
'Terceira Banda - 3 7750610011 2
'3 Dígito verificador da primeira banda
'7750610011 Número da Conta
'2 Dígito verificador da terceira banda
'-------------------------------------------

    cmdLimparCampos_Click
    cboBanco.ListIndex = 20    '237 - BANCO BRADESCO S.A.
    txtAgencia.Text = "0494"
    txtComp.Text = "018"
    txtChequeNumero.Text = "001793"
    cboTipificacaoCheque.ListIndex = 2  '5 - Comum
    txtConta.Text = "7750610011"
    cmdCalcular_Click
End Sub

Private Sub cmdExemplo3_Click()
'-------------------------------------------
'Primeira Banda
'001 - Código do Banco
'0005 - Código da Agência
'8 - Dígito Verificador da Segunda Banda
'-------------------------------------------
'Segunda Banda
'001 - Código da Câmara de Compensação
'218917 - Número do Cheque
'5 - Tipificação do Cheque (5–Comum 6–Bancário 7–Salário 8–Administrativo 9–CPMF)
'-------------------------------------------
'Terceira Banda
'7- Dígito verificador da Primeira Banda
'0400168645 - Número da Conta Corrente
'2 - Dígito verificador da Terceira Banda
'-------------------------------------------

    cmdLimparCampos_Click
    cboBanco.ListIndex = 38    '001 - BANCO DO BRASIL S.A.
    txtAgencia.Text = "0005"
    txtComp.Text = "001"
    txtChequeNumero.Text = "218917"
    cboTipificacaoCheque.ListIndex = 2  '5 - Comum
    txtConta.Text = "0400168645"
    cmdCalcular_Click
End Sub
