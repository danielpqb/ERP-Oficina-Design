Attribute VB_Name = "NewInput"
Option Explicit

Dim i               As Integer

Public SQL2         As String  'Comando SQL que será utilizado para gravar cada linha (armazena os valores de cada coluna)
Public lcab         As Integer 'Ajuda a definir o 'nomecoluna'
Public ln           As Integer 'Linha de novos dados

Sub InputBaseNewSheet()

'valor001 = Produtos
'valor002 = Tb
'BaseNewSheet = TbProdutos

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "valor001"
tipotabela = "valor002"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.BaseNewSheet
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.BaseNewSheet
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.BaseNewSheet
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.BaseNewSheet
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListBaseNewSheet

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputTbProdutos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Produtos"
tipotabela = "Tb"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.TbProdutos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.TbProdutos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.TbProdutos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.TbProdutos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListTbProdutos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsPartes_Produtos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Partes_Produtos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsPartes_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsPartes_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsPartes_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsPartes_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsPartes_Produtos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsClientes()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Clientes"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    If ActiveSheet.optButton1CnsClientes = True Then
        SQL = "INSERT INTO " & tipotabela & nomedasheet & "PF("
    ElseIf ActiveSheet.optButton2CnsClientes = True Then
        SQL = "INSERT INTO " & tipotabela & nomedasheet & "PJ("
    End If
    
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsClientes
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsClientes
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsClientes
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsClientes
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsClientes

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsPedidos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Pedidos"
tipotabela = "Cns"

numpedido = 0 'Necessario para calcular o numero do pedido na Validacao de Dados

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsPedidos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsPedidos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsPedidos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsPedidos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsPedidos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsInsumos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Insumos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsInsumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsInsumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsInsumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsInsumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsInsumos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsInsumos_Produtos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Insumos_Produtos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsInsumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsInsumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsInsumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsInsumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsInsumos_Produtos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsCortes_Insumos_Produtos()

'Cortes_Insumos_Produtos = Produtos
'Cns = Tb
'CnsCortes_Insumos_Produtos = TbProdutos

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Cortes_Insumos_Produtos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsCortes_Insumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsCortes_Insumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsCortes_Insumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsCortes_Insumos_Produtos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsCortes_Insumos_Produtos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub

Sub InputCnsCompras_Insumos()

'Compras_Insumos = Produtos
'Cns = Tb
'CnsCompras_Insumos = TbProdutos

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

nomedasheet = "Compras_Insumos"
tipotabela = "Cns"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

lcab = Range("New" & tipotabela & nomedasheet).Row          'Linha do cabeçalho de novos dados

ninputedit = 1 'Indica para a validação de dados que é um input de 'Novo Registro'
cor = "" 'Indica para a Validação de Dados a Cor da célula

For i = 1 To ActiveSheet.txtboxQntNewRows.Value

    'Capturar novos valores para as variáveis
    '-------------------------------------------------
    ln = Range("New" & tipotabela & nomedasheet).Row + i    'Linha de novos dados
    
    'Define o comando SQL, e calcula e valida os dados inseridos
    '-----------------------------------------------------------
    SQL = "INSERT INTO " & tipotabela & nomedasheet & "("
    SQL2 = ""
    
    If Range("A" & ln).Value = "MODIFICADO" And Range("A" & ln).Interior.Color = RGB(255, 140, 50) Then
    
        For Each coluna In Range("New" & tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(lcab, coluna.Column)
            
            ''BRANCO e AZUL
            If Cells(ln, coluna.Column).Interior.Color = RGB(240, 240, 240) Or Cells(ln, coluna.Column).Interior.Color = RGB(215, 225, 245) Then
                SQL = SQL & nomecoluna & ", "
                SQL2 = SQL2 & "'" & Cells(ln, coluna.Column) & "', "
            
            ''PRETO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                cor = RGB(90, 90, 90)
                valorpreto = ""
                Call ValidateData.CnsCompras_Insumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorpreto & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorpreto & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula PRETA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
            
            ''AMARELO
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                cor = RGB(255, 240, 205)
                valoramarelo = Cells(ln, coluna.Column)
                Call ValidateData.CnsCompras_Insumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valoramarelo & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valoramarelo & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
             ''LARANJA
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                cor = RGB(255, 230, 205)
                valorlaranja = ""
                Call ValidateData.CnsCompras_Insumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorlaranja & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorlaranja & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            ''VERDE
            ElseIf Cells(ln, coluna.Column).Interior.Color = RGB(225, 240, 220) Then
                cor = RGB(225, 240, 220)
                valorverde = Cells(ln, coluna.Column)
                Call ValidateData.CnsCompras_Insumos
                If existe = True Then
                    If bloqueado = True Then
                        MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não permite o valor '" & valorverde & "'!")
                        End
                    ElseIf bloqueado = False Then
                        If ignorado = False Then
                            SQL = SQL & nomecoluna & ", "
                            SQL2 = SQL2 & "'" & valorverde & "', "
                        End If
                    End If
                ElseIf existe = False Then
                    MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula VERDE tenha uma validação de dados quando se cria um novo registro!")
                    cn.Close
                    End
                End If
                
            Else
                MsgBox ("A célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' possui uma cor não especificada!")
                cn.Close
                End
            End If
    
        Next
        
        SQL = Left(SQL, Len(SQL) - 2) & ") VALUES (" & Left(SQL2, Len(SQL2) - 2) & ");"
        
        ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
        ''---------------------------------------
        Call General.OpenRS 'Por algum motivo, ao usar um INSERT INTO, não preciso dar 'close' no 'recordset'
        
        ''LIMPA A LINHA INCLUÍDA COMO NOVO REGISTRO
        ''-----------------------------------------
        Rows(Range("A" & ln).Row).ClearContents
        Range("A" & ln).ClearFormats
    
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA
''-------------------------------------
Call ListData.ListCnsCompras_Insumos

Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
