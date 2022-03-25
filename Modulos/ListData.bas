Attribute VB_Name = "ListData"
'Outputs
'''''''''''''''''''''''''''''''''''''''''''''
Dim filtrou         As Integer              'Diz se a lista foi filtrada com o comando WHERE
'''''''''''''''''''''''''''''''''''''''''''''

'InserirDados e FormatarDados
'''''''''''''''''''''''''''''''''''''''''''''
Dim newlin          As Integer              'Linha do cabeçalho de inserção de novos dados
Dim col             As Integer              'Coluna de início do cabeçalho
Dim ccol            As Integer              'Contador para inserir cabeçalho
'''''''''''''''''''''''''''''''''''''''''''''

Sub ListBaseNewSheet()

'valor001 = Produtos
'valor002 = Tb
'valor003 = PKProduto
'BaseNewSheet = TbProdutos

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "valor001"
tipotabela = "valor002"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE valor003 <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY valor003;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Private Sub InsertData()

'Declara as Variáveis utilizadas
'-------------------------------------------
Dim FD              As ADODB.Field          'Campo do Banco de Dados
Dim lin             As Integer              'Linha do cabeçalho
'-------------------------------------------

'Primeira linha e coluna do cabeçalho
'---------------------------------------------
col = ActiveSheet.Range(tipotabela & nomedasheet).Column
lin = ActiveSheet.Range(tipotabela & nomedasheet).Row
ccol = col 'Contador de colunas do cabeçalho
newlin = ActiveSheet.Range("New" & tipotabela & nomedasheet).Row

'Apaga as células utilizadas anteriormente
'---------------------------------------------
With ActiveSheet
    .Range("A:A", "1:1").ClearFormats
    .Range("A:A", "1:1").ClearContents
    .Range(tipotabela & nomedasheet).ClearFormats
    .Range(tipotabela & nomedasheet).ClearContents
    .Range(newlin & ":" & lin - 2).ClearFormats
    .Range(newlin & ":" & lin - 2).ClearContents
End With
    
'Adicionar o nome das colunas
'-------------------------------------------
For Each FD In rs.Fields

    'Cabeçalho da tabela
    With ActiveSheet.Cells(lin, ccol)
        .Value = FD.Name
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'Cabeçalho da inserção de novos dados
    With ActiveSheet.Cells(newlin, ccol)
        .Value = FD.Name
        .Font.Bold = True
        .Interior.Color = RGB(38, 38, 38)
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
            
    ccol = ccol + 1

Next FD

'Verifica se há dados no Recordset
'-----------------------------------------------
If rs.EOF = False Then  'EOF = End of File
    
    'Inserir dados do Recordset na planilha
    '------------------------------------------
    ActiveSheet.Cells(lin + 1, col).CopyFromRecordset rs
    
Else
 
    MsgBox "Não há dados para serem trazidos..."
    
End If

End Sub

Private Sub FormatData()

'Declara as Variáveis utilizadas
'-------------------------------------------
Dim cornewcol       As Integer              'Cor da coluna de inserção de dados
Dim arrbranco       As Variant              'Array das células que serão coloridas de azul
Dim arramarelo      As Variant              'Array das células que serão coloridas de verde
Dim arrverde        As Variant              'Array das células que serão coloridas de verde
Dim valarr          As Variant              'Valor que percorre os arrays
Dim ocultarcolunas  As Variant              'Array das colunas que serão ocultadas

Dim corcb As Variant
corcb = RGB(230, 230, 230)
Dim valcb As Variant
Dim colesp As Integer
colesp = 0
Dim colunai As Integer
Dim colunaf As Integer
Dim linhai As Integer
Dim linhaf As Integer
'-------------------------------------------

'Formata as células da tabela
'---------------------------------------------
With ActiveSheet.Range(tipotabela & nomedasheet)
    .CurrentRegion.Borders.LineStyle = xlContinuous
    .CurrentRegion.HorizontalAlignment = xlCenter
    .CurrentRegion.VerticalAlignment = xlCenter
    .CurrentRegion.AutoFilter
    .CurrentRegion.Columns.AutoFit
End With

'Formata as células da tabela de inserção de novos dados
'-------------------------------------------------------
With ActiveSheet.Range(Cells(newlin + 1, col), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, ccol - 1))
    .ClearFormats
    .Borders.LineStyle = xlContinuous
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Color = RGB(0, 0, 0)
    .Font.Bold = True
End With

Dim tbnewcol As Range 'Contador de colunas

''Define e colore as colunas com suas respectivas cores
''-----------------------------------------------------
For Each tbnewcol In Range(Cells(newlin, col), Cells(newlin, ccol - 1))

    cornewcol = 0
    
    ''Define os arrays das tabelas, separando por cores, e por colunas a serem ocultadas
    ''Consulte a Worksheet 'Tabela de Cores' para mais informações
    'Preto = 1
    'Branco = 2
    'Amarelo = 3
    'Laranja = 4
    'Verde = 5
    'Azul = 6
        
    'TbProdutos
    If tipotabela & nomedasheet = "TbProdutos" Then
        arrbranco = Array("Descricao_Produto", "Largura_Produto", "Altura_Produto", "Profundidade_Produto", "FotosProjetos_Produto", "Preco_Produto", "InfoFerro_Produto", "InfoMadeira_Produto", "InfoVidroAcrilico_Produto", "InfoTecido_Produto", "Info_Produto")
        arramarelo = Array("Categoria_Produto", "Linha_Produto", "Tipo_Produto", "SR", "MC", "VD", "ES")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsPartes_Produtos
    ElseIf tipotabela & nomedasheet = "CnsPartes_Produtos" Then
        arrbranco = Array("Descricao_Parte_Produto")
        arramarelo = Array("FKProduto")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsClientes
    ElseIf tipotabela & nomedasheet = "CnsClientes" Then
        'PF
        If ActiveSheet.optButton1CnsClientes = True Then
            arrbranco = Array("Nome_PF", "CPF", "RG", "Email", "Telefone")
            arramarelo = Array("")
            arrlaranja = Array("")
            arrverde = Array("")
            arrazul = Array("")
            ocultarcolunas = Array("")
        'PJ
        ElseIf ActiveSheet.optButton2CnsClientes = True Then
            arrbranco = Array("NomeFantasia_PJ", "CNPJ", "Email", "Telefone")
            arramarelo = Array("")
            arrlaranja = Array("")
            arrverde = Array("")
            arrazul = Array("")
            ocultarcolunas = Array("")
        End If
    'CnsPedidos
    ElseIf tipotabela & nomedasheet = "CnsPedidos" Then
        arrbranco = Array("Valor_Venda")
        arramarelo = Array("FKCliente", "FKProduto", "Quantidade", "Data_Venda")
        arrlaranja = Array("QntProducao", "Cancelado", "Data_Entrega")
        arrverde = Array("Data_LimiteEntrega")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsInsumos
    ElseIf tipotabela & nomedasheet = "CnsInsumos" Then
        arrbranco = Array("UnidadePorPeso", "Categoria_Insumo", "Descricao_Insumo", "UnCompra_Insumo", "UnMed_Insumo", "Perda_Insumo", "Preco_UnCompra", "UnMedPorCompra")
        arramarelo = Array("")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsInsumos_Produtos
    ElseIf tipotabela & nomedasheet = "CnsInsumos_Produtos" Then
        arrbranco = Array("")
        arramarelo = Array("FKInsumo", "FKProduto", "Qnt")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsCortes_Insumos_Produtos
    ElseIf tipotabela & nomedasheet = "CnsCortes_Insumos_Produtos" Then
        arrbranco = Array("Descricao_Corte")
        arramarelo = Array("FKInsumo_Produto", "Qnt")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsProducao_Pedidos
    ElseIf tipotabela & nomedasheet = "CnsProducao_Pedidos" Then
        arrbranco = Array("Inicio_SR", "Inicio_SR", "Fim_SR", "Inicio_MC", "Fim_MC", "Inicio_PT_SR", "Fim_PT_SR", "Inicio_PT_MC", "Fim_PT_MC", "Inicio_VD", "Fim_VD", "Inicio_ES", "Fim_ES", "MontarConferirEmbalar")
        arramarelo = Array("Material", "Envio", "Plano_Corte")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsCompras_Insumos
    ElseIf tipotabela & nomedasheet = "CnsCompras_Insumos" Then
        arrbranco = Array("Data_Compra")
        arramarelo = Array("FKInsumo", "FKFornecedor", "Qnt", "Preco")
        arrlaranja = Array("Recebido")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    'CnsEstoque_Insumos
    ElseIf tipotabela & nomedasheet = "CnsEstoque_Insumos" Then
        arrbranco = Array("")
        arramarelo = Array("")
        arrlaranja = Array("Contagem")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    Else 'Não especificado
        arrbranco = Array("")
        arramarelo = Array("")
        arrlaranja = Array("")
        arrverde = Array("")
        arrazul = Array("")
        ocultarcolunas = Array("")
    End If
    
    'Define as cores de cada planilha
    For valarr = LBound(arrbranco) To UBound(arrbranco)
        If arrbranco(valarr) = tbnewcol.Value Then
            cornewcol = 2
            GoTo colorir
        End If
    Next
    For valarr = LBound(arramarelo) To UBound(arramarelo)
        If arramarelo(valarr) = tbnewcol.Value Then
            cornewcol = 3
            GoTo colorir
        End If
    Next
    For valarr = LBound(arrlaranja) To UBound(arrlaranja)
        If arrlaranja(valarr) = tbnewcol.Value Then
            cornewcol = 4
            GoTo colorir
        End If
    Next
    For valarr = LBound(arrverde) To UBound(arrverde)
        If arrverde(valarr) = tbnewcol.Value Then
            cornewcol = 5
            GoTo colorir
        End If
    Next
    For valarr = LBound(arrazul) To UBound(arrazul)
        If arrazul(valarr) = tbnewcol.Value Then
            cornewcol = 6
            GoTo colorir
        End If
    Next
    '-----------------------------------------
    
colorir:
    'Colore as colunas com as cores definidas
    If cornewcol = 2 Then 'Branco
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Interior.Color = RGB(240, 240, 240)
    ElseIf cornewcol = 3 Then 'Amarelo
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Interior.Color = RGB(255, 240, 205)
    ElseIf cornewcol = 4 Then 'Laranja
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Interior.Color = RGB(255, 230, 205)
    ElseIf cornewcol = 5 Then 'Verde
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Interior.Color = RGB(225, 240, 220)
    ElseIf cornewcol = 6 Then 'Azul
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Interior.Color = RGB(215, 225, 245)
    Else 'Preto
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Interior.Color = RGB(90, 90, 90)
        Range(Cells(newlin + 1, tbnewcol.Column), Cells(newlin + ActiveSheet.txtboxQntNewRows.Value, tbnewcol.Column)).Font.Color = RGB(255, 255, 255)
    '------------------------------------------
    End If
    
    'Oculta as colunas escolhidas
    For valarr = LBound(ocultarcolunas) To UBound(ocultarcolunas)
        If ocultarcolunas(valarr) = tbnewcol.Value Then
            tbnewcol.EntireColumn.Hidden = True
        End If
    Next
    
    'Cinza/Branco
    If nomecolesp = tbnewcol.Value Then
        colesp = tbnewcol.Column
    End If
    
Next tbnewcol

'Redefine os NOMES das células
'---------------------------------------------
ThisWorkbook.Names.Add Name:=tipotabela & nomedasheet, _
                       RefersTo:=ActiveSheet.Range(tipotabela & nomedasheet).CurrentRegion, _
                       Visible:=True
ThisWorkbook.Names.Add Name:="New" & tipotabela & nomedasheet, _
                       RefersTo:=ActiveSheet.Range("New" & tipotabela & nomedasheet).CurrentRegion, _
                       Visible:=True
                       
'Agrupa as linhas separando nas cores Cinza e Branco
'-------------------------------------------------------------------
colunai = Range(tipotabela & nomedasheet).Column
colunaf = Range(tipotabela & nomedasheet).Columns.End(xlToRight).Column
If Range(tipotabela & nomedasheet).Rows.End(xlDown).Row < 100000 Then
    linhaf = Range(tipotabela & nomedasheet).Rows.End(xlDown).Row
Else
    linhaf = Range(tipotabela & nomedasheet).Row + 1
End If

If colesp = 0 Then
    colesp = colunai
End If

linhai = Range(tipotabela & nomedasheet).Row + 1
'valcb = Cells(linhai, colesp).Value
corcb = RGB(0, 0, 0)

For i = linhai To linhaf - 1
    If Cells(i, colesp).Value <> Cells(i + 1, colesp).Value Then
        If corcb = RGB(0, 0, 0) Then
            corcb = RGB(242, 242, 248)
            Range(Cells(i + 1, colunai).Address, Cells(i + 1, colunaf).Address).Interior.Color = corcb
        Else
            corcb = RGB(0, 0, 0)
        End If
    Else
        If corcb = RGB(242, 242, 248) Then
            Range(Cells(i + 1, colunai).Address, Cells(i + 1, colunaf).Address).Interior.Color = corcb
        End If
    End If
Next


End Sub

Sub ListTbProdutos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Produtos"
tipotabela = "Tb"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKProduto <> 1 "

'Filtros com CheckBox
filtrou = 0

If (ActiveSheet.chkbox1TbProdutos = True And ActiveSheet.chkbox2TbProdutos = True And ActiveSheet.chkbox3TbProdutos = True) Then
    GoTo ordenardados
ElseIf (ActiveSheet.chkbox1TbProdutos = False And ActiveSheet.chkbox2TbProdutos = False And ActiveSheet.chkbox3TbProdutos = False) Then
    ActiveSheet.chkbox1TbProdutos = True
    ActiveSheet.chkbox2TbProdutos = True
    ActiveSheet.chkbox3TbProdutos = True
    GoTo ordenardados
Else
    SQL = SQL & " AND (Tipo_Produto = "
    If ActiveSheet.chkbox1TbProdutos = True Then
        SQL = SQL & "'EL'"
        filtrou = 1
    End If
    If ActiveSheet.chkbox2TbProdutos = True Then
        If filtrou = 0 Then
            SQL = SQL & "'ESP'"
            filtrou = 1
        Else
            SQL = SQL & " OR Tipo_Produto = 'ESP'"
        End If
    End If
    If ActiveSheet.chkbox3TbProdutos = True Then
        If filtrou = 0 Then
            SQL = SQL & "'FL'"
            filtrou = 1
        Else
            SQL = SQL & " OR Tipo_Produto = 'FL'"
        End If
    End If
    SQL = SQL & ")"
End If

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY PKProduto;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsPartes_Produtos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Partes_Produtos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKParte_Produto <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY PKParte_Produto;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
nomecolesp = "FKProduto"
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsClientes()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Clientes"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
If ActiveSheet.optButton1CnsClientes = True Then
    SQL = "SELECT * FROM " & tipotabela & nomedasheet & "PF WHERE PKCliente <> 1 "
ElseIf ActiveSheet.optButton2CnsClientes = True Then
    SQL = "SELECT * FROM " & tipotabela & nomedasheet & "PJ WHERE PKCliente <> 1 "
End If

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY PKCliente;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsPedidos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Pedidos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKPedido <> 1 "

If ActiveSheet.chkbox4 = True Then
    SQL = SQL & "AND (Quantidade <> QntProducao) "
    ActiveSheet.chkbox1 = False
    ActiveSheet.chkbox2 = False
    ActiveSheet.chkbox3 = False
    GoTo ordenardados
End If

If ActiveSheet.chkbox1 = False Then
    If ActiveSheet.chkbox2 = False And ActiveSheet.chkbox3 = False Then
        'FFF
        ActiveSheet.chkbox1 = True
        SQL = SQL & "AND (Data_Entrega = '' OR Data_Entrega IS NULL) AND (Cancelado = '' OR Cancelado IS NULL) "
    ElseIf ActiveSheet.chkbox2 = False And ActiveSheet.chkbox3 = True Then
        'FFV
        SQL = SQL & "AND Cancelado <> '' "
    ElseIf ActiveSheet.chkbox2 = True And ActiveSheet.chkbox3 = False Then
        'FVF
        SQL = SQL & "AND Data_Entrega <> '' "
    ElseIf ActiveSheet.chkbox2 = True And ActiveSheet.chkbox3 = True Then
        'FVV
        SQL = SQL & "AND Data_Entrega <> '' OR Cancelado <> '' "
    End If
    
ElseIf ActiveSheet.chkbox2 = False And ActiveSheet.chkbox3 = False Then
    'VFF
    SQL = SQL & "AND (Data_Entrega = '' OR Data_Entrega IS NULL) AND (Cancelado = '' OR Cancelado IS NULL) "
ElseIf ActiveSheet.chkbox2 = False And ActiveSheet.chkbox3 = True Then
    'VFV
    SQL = SQL & "AND (Data_Entrega = '' OR Data_Entrega IS NULL) OR Cancelado <> '' "
ElseIf ActiveSheet.chkbox2 = True And ActiveSheet.chkbox3 = False Then
    'VVF
    SQL = SQL & "AND (Cancelado = '' OR Cancelado IS NULL) "
ElseIf ActiveSheet.chkbox2 = True And ActiveSheet.chkbox3 = True Then
    'VVV
    'Nao faz nada
End If

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY PKPedido;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
nomecolesp = "Numero_Pedido"
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsInsumos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Insumos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKInsumo <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY Descricao_Insumo;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
nomecolesp = "Categoria_Insumo"
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsInsumos_Produtos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Insumos_Produtos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKInsumo_Produto <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY FKProduto, Descricao_Insumo;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
nomecolesp = "FKProduto"
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub


Sub ListCnsCortes_Insumos_Produtos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Cortes_Insumos_Produtos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKCorte_Insumo_Produto <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY Cod_Produto, Descricao_Insumo, Descricao_Corte;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
nomecolesp = "Cod_Produto"
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsProducao_Pedidos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Producao_Pedidos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKProducao_Pedido <> 1 "

'Filtros com CheckBox
filtrou = 0

If (ActiveSheet.chkbox1 = True And ActiveSheet.chkbox2 = True) Then
    GoTo ordenardados
ElseIf (ActiveSheet.chkbox1 = False And ActiveSheet.chkbox2 = False) Then
    ActiveSheet.chkbox1 = True
    ActiveSheet.chkbox2 = True
    GoTo ordenardados
Else
    SQL = SQL & " AND (Envio "
    If ActiveSheet.chkbox1 = True Then
        SQL = SQL & "IS NULL OR Envio = ''"
        filtrou = 1
    End If
    If ActiveSheet.chkbox2 = True Then
        If filtrou = 0 Then
            SQL = SQL & "<> ''"
            filtrou = 1
        Else
            SQL = SQL & " OR Envio <> ''"
        End If
    End If
    SQL = SQL & ")"
End If

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY PKProducao_Pedido;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsCompras_Insumos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Compras_Insumos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKCompra_Insumo <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY PKCompra_Insumo;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub

Sub ListCnsEstoque_Insumos()

Application.EnableEvents = False
Application.ScreenUpdating = False

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Estoque_Insumos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

'Define o comando a ser utilizado no Banco de Dados (Linguagem SQL)
'------------------------------------------------------------------
SQL = "SELECT * FROM " & tipotabela & nomedasheet & " WHERE PKInsumo <> 1 "

ordenardados:
'Ordena os dados
SQL = SQL & " ORDER BY Categoria_Insumo, Descricao_Insumo;"

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
Call General.OpenRS

''INSERE OS DADOS OBTIDOS NA PLANILHA
''-----------------------------------
Call ListData.InsertData

''FECHA A CONSULTA FEITA NO BANCO DE DADOS
''----------------------------------------
Call General.CloseRS

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''FORMATA AS CÉLULAS DA PLANILHA
''------------------------------
nomecolesp = "Categoria_Insumo"
Call ListData.FormatData

'PERMITE EDITAR ALGUNS INTERVALOS PROTEGIDOS
'-------------------------------------------
Call General.AllowProtectEdit

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.StatusBar = "Listagem atualizada em " & Now()

End Sub
