Attribute VB_Name = "General"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Usado em ListData, EditData, NewInput e ValidateData
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public cn            As New ADODB.Connection 'Conecta ao Banco de Dados
Public rs            As New ADODB.Recordset  'Executa um comando SQL no Banco de Dados
Public rs2          As New ADODB.Recordset
Public arq           As String               'Caminho do Banco de Dados
Public conectadb     As String               'String com comando e caminho necessários para conectar ao Banco de Dados
Public SQL           As String               'Comando SQL a ser usado no Banco de Dados

'''''''''''''''''''''''''''''''''''''''
'Usado em ListData, EditData e NewInput
'''''''''''''''''''''''''''''''''''''''
Public nomedasheet   As String               'Nome da Sheet utilizada
Public tipotabela    As String               'Tipo de tabela do Banco de Dados (Cns ou Tb)

'''''''''''''''''''''''''''''''''''''''''''
'Usado em EditData, NewInput e ValidateData
'''''''''''''''''''''''''''''''''''''''''''
'Geral
Public valorpreto       As String
Public valoramarelo     As String
Public valorlaranja     As String
Public valorverde       As String
Public nomecoluna       As String
Public ignorado         As Boolean
Public bloqueado        As Boolean
Public existe           As Boolean
Public ninputedit       As Integer 'Validação de Novo Registro = 1 ; Validação de Edit = 2
Public cor              As Variant 'Indica para a Validação de Dados a Cor da célula
Public coluna           As Variant
Public nomecolesp As String 'Nome da coluna que especifica os agrupamentos das linhas nas cores Cinza RGB(230,230,230) e Branco (inalterado)
Public pk                   As Integer 'Salva o numero do PK do registro trabalhado, para que posso ser usado em outros blocos para procura. (Necessario para fazer insercao de pedidos em producao)
'CnsClientes
Public novopkcliente    As Integer
Public fkpessoafisica   As Integer
Public fkpessoajuridica As Integer
'CnsPedidos
Public numpedido As Integer

Sub DeclarePublic()

'Quando esse Sub é chamado as variáveis desse módulo são declaradas

End Sub

Sub DefineDBConection()

arq = ThisWorkbook.Path & "\DB.accdb"

conectadb = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & arq & ";Persist Security Info=False"

''Define uma nova conexão com o Banco de Dados
''---------------------------------------------
'Set cn = New ADODB.Connection

End Sub

Sub ConectDB()

'Abre a conexão de dados
'----------------------------------------------
cn.Open conectadb

End Sub

Sub OpenRS()

''Define um novo recordset
''----------------------------------------------
'Set rs = New ADODB.Recordset

'Realiza a consulta
'-----------------------------------------------
rs.Open SQL, cn

End Sub

Sub CloseRS()

rs.Close

End Sub

Sub PasteValues()
Attribute PasteValues.VB_ProcData.VB_Invoke_Func = "v\n14"

''Esse processo está vinculado ao atalho 'CTRL + V'

On Error GoTo fim

Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Exit Sub

fim:
MsgBox "Faça apenas colagem simples para evitar bugs e crashes na planilha!"

End Sub

Sub DisconectDB()

cn.Close

End Sub

Sub ProtectSheet()

'If Application.UserName <> "Daniel" Then
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowInsertingHyperlinks:=True, Password:="123", AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
'End If

End Sub

Sub UnprotectSheet()

ActiveSheet.Unprotect Password:="123"

End Sub

Sub AllowProtectEdit()

Dim pc As Integer
Dim nc As Integer

Dim prnt As Integer
Dim nrnt As Integer

Dim prt As Integer
Dim nrt As Integer

Dim rnge As String
Dim titl As String

titl = tipotabela & nomedasheet

pc = ActiveSheet.Range(titl).Column
nc = ActiveSheet.Range(titl).Columns.Count

prnt = ActiveSheet.Range("New" & titl).Row
nrnt = ActiveSheet.txtboxQntNewRows.Value

prt = ActiveSheet.Range(titl).Row
nrt = ActiveSheet.Range(titl).Rows.Count

rnge = Cells(prt + 1, pc + 1).Address(0, 0) & ":" & Cells(nrt + prt - 1, nc + pc - 1).Address(0, 0) & "," & Cells(prnt + 1, pc).Address(0, 0) & ":" & Cells(nrnt + prnt, nc + pc - 1).Address(0, 0)

ActiveSheet.Cells.Select
Selection.Locked = True

ActiveSheet.Range("A1").Select

On Error Resume Next
ActiveSheet.Protection.AllowEditRanges(titl).Delete
ActiveSheet.Protection.AllowEditRanges.Add Title:=titl, Range:=Range(rnge)

End Sub
