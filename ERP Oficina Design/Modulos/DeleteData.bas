Attribute VB_Name = "DeleteData"
Sub DeleteDataBaseNewSheet()

'valor001 = Produtos
'valor002 = Tb
'valor003 = PKProduto
'BaseNewSheet = TbProdutos

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "valor002"
nomedasheet = "valor001"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " valor003 = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub

Sub DeleteDataTbProdutos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Tb"
nomedasheet = "Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKProduto = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub

Sub DeleteDataCnsPartes_Produtos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Partes_Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKParte_Produto = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub

Sub DeleteDataCnsClientes()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Clientes"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
If ActiveSheet.optButton1CnsClientes = True Then
    SQL = "DELETE * FROM " & tipotabela & nomedasheet & "PF WHERE"
ElseIf ActiveSheet.optButton2CnsClientes = True Then
    SQL = "DELETE * FROM " & tipotabela & nomedasheet & "PJ WHERE"
End If

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKCliente = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub

Sub DeleteDataCnsPedidos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Pedidos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKPedido = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub


Sub DeleteDataCnsInsumos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Insumos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKInsumo = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub


Sub DeleteDataCnsInsumos_Produtos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Insumos_Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKInsumo_Produto = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub

Sub DeleteDataCnsCortes_Insumos_Produtos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Cortes_Insumos_Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKCorte_Insumo_Produto = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub

Sub DeleteDataCnsCompras_Insumos()

Application.ScreenUpdating = False

''DESPROTEGE A PLANILHA
''---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Compras_Insumos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

'Define o comando SQL
'----------------------------------------------
SQL = "DELETE * FROM " & "Tb" & nomedasheet & " WHERE"

''DEFINE 'delselect' COMO O INTERVALO DA PRIMEIRA COLUNA DA TABELA, IGNORANDO O CABEÇALHO
''---------------------------------------------------------------------------------------
Set delselect = Intersect(Selection.Rows, Rows(Range(tipotabela & nomedasheet).Row + 1 & ":" & (Range(tipotabela & nomedasheet).Rows.Count + Range(tipotabela & nomedasheet).Row - 1)), Columns(Range(tipotabela & nomedasheet).Column))

If delselect Is Nothing Then
    MsgBox "Nenhuma linha selecionada!"
    GoTo fim
Else
    yn = MsgBox("Tem certeza que deseja excluir as linhas selecionadas?" & vbCrLf & "Após excluir os dados, não será mais possível recuperá-los!", vbYesNo, "Alerta!")
    If yn = vbYes Then
        For Each linha In delselect.Rows
            SQL = SQL & " PKCompra_Insumo = " & Cells(linha.Row, linha.Column) & " OR"
        Next
        SQL = Left(SQL, Len(SQL) - 3) & ";"
    Else 'Se a resposta for não
        Call General.ProtectSheet
        End
    End If
End If

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

''EXECUTA O COMANDO SQL NO BANCO DE DADOS
''---------------------------------------
cn.Execute SQL

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Application.Run ("ListData.List" & tipotabela & nomedasheet)

fim:
Application.ScreenUpdating = True

End Sub
