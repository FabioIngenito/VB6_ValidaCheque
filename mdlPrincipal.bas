Attribute VB_Name = "mdlPrincipal"
Option Explicit
Public strPath As String
Public strSQL As String

'Dimensiona a Conex�o ao Banco de Dados
Public adoCMC7 As ADODB.Connection

'Dimensiona as "Class Module"
Public BD As clsBD
Public modul As clsModulo

'Dimensiona e Instancia a "Class Module"
Public Geral As New clsGeral

Sub Main()
    strPath = App.Path & "\"
   
    frmValidaCheque.Show
End Sub

'=====================================================
'http://www.veloso.adm.br/artigocmc7.asp
'=====================================================

'D�vidas sobre cadastramento de cheques em programas

'O que � Linha1 do Cheque?
'� a parte superior do cheque onde tem os dados, na ordem: _
 |Comp | Banco | ag�ncia |C1| Conta |C2| |Cheque|C3| R$ |

'Para que serve a linha1?
'Serve para digita��o do cheque em sistemas quando n�o for _
 poss�vel digitar o CMC7.

'O que � CMC7 do Cheque?
'CMC7 significa Caracteres Magn�ticos Codificados em 7 barras.
'Tamb�m � chamada de banda magn�tica ou linha2 do cheque. _
 � apenas mais um padr�o de c�digo de barras, por�m com uma _
 particularidade, a impress�o tem que ser com toner magn�tico.
'Al�m da impress�o com toner magn�tico, outra forma de dificultar _
 a falsifica��o � a valida��o dos 3 d�gitos verificadores _
 contidos no CMC7.

'Os separadores dividem o CMC7 em tr�s bandas.
'Ex:
'Primeira Banda Segunda Banda Terceira Banda
'23704948       0180017935    377506100112
 
'A valida��o do CMC7 � padr�o para todos os bancos?
'Sim. N�o h� uma nova regra para valida��o do CMC7, tamb�m a _
 valida��o n�o depende do banco, esta regra � da Febraban. _
 A altera��o desta regra afetaria muitos sistemas.
'Por Favor, n�o telefone ou envie e-mail perguntando sobre _
altera��o.

'Porque o n�mero da conta alguns de bancos � diferente na Linha1 _
 e CMC7 ?
'Isto ocorre porque alguns bancos aproveitam algumas posi��es da _
 conta do CMC7 para colocarem n�meros que n�o teriam espa�o no _
 CMC7, estes n�meros n�o est�o na conta da Linha1.
'Estes n�meros s�o chamados de 'Raz�o' da conta.

'Os campos C1, C2 e C3 est�o no CMC7 ?
'N�o, os campos C1, C2 e C3 s�o d�gitos de controle para digita��o _
 da linha1. Os d�gitos de controle do CMC7 tamb�m s�o 3 por�m _
 calculados de forma diferente.

'Porque o componente TCMC7 retorna a conta do banco diferente da _
 linha1 do cheque ?
'O objetivo principal do componente TCMC7 � validar o CMC7 e a _
 Linha1, ele faz isto atrav�s dos 3 d�gitos da Linha1 (C1,C2, e C3) _
 e dos 3 d�gitos do CMC7.
'O componente tem uma fun��o para retornar a conta do CMC7, por�m _
 sempre retornar� 10 caracteres referentes ao n�mero da conta do CMC7.
'Isto n�o impede que seja feita um fun��o para tratar este n�mero _
 de conta, isto n�o � feito no componente porque n�o existe uma _
 regra padronizada.

'Quando fizer o cadastramento do cheque atrav�s do CMC7 tenho que _
 mudar o n�mero da conta ?
'Se for gerar arquivo de compensa��o para banco tipo CNAB240 n�o _
 h� necessidade, pois neste tipo de arquivo voc� fornece o CMC7 _
 inteiro com foi capturado ou a Linha1 como foi digitada.
'Agora para guardar na sua base de dados, o melhor seria mudar o _
 n�mero da conta para bancos que utilizam a RAZ�O da conta.

'exemplo:  Ita�
'Conta na Linha1: 23288-2
'Conta no CMC7: 7123232882
'(7123=n�mero que muda para cada folha do cheque)
'(232882 = n�mero da conta)

'Ent�o, se voc� utilizar o n�mero da conta para buscar dados em _
 outras tabelas, n�o vai funcionar pois o n�mero sempre ser� _
 diferente.
'Neste caso voc� deveria retirar o 7123, 4 primeiras posi��es, _
 antes de fazer a pesquisa.

'� poss�vel transformar a linha1 em CMC7 e vice-versa ?
'N�o. Os Bancos criaram o campo raz�o para dificultar isto. _
 O campo raz�o esta marcado com aster�sco no CMC7 abaixo, varia _
 de tamanho conforme o banco. A regra para cria��o da raz�o n�o _
 existe padroniza��o ou legisla��o e portanto � um segredo de _
 cada banco.
