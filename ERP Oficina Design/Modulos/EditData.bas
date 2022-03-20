Attribute VB_Name = "EditData"
Option Explicit

Dim editou                  As Boolean
Dim editouqlqcoisa          As Boolean
Dim arrignocel              As Variant
Dim arrignoval              As Variant
Dim vai                     As Variant 'Contador que percorre o Array dos nomes e valores das células ignoradas

Public linha                As Variant
Public msgignorados         As String

Sub EditBaseNewSheet()

'valor001 = Produtos
'valor002 = Tb
'valor003 = PKProduto
'BaseNewSheet = TbProdutos

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "valor002"
nomedasheet = "valor001"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Or Left(Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column), 2) = "PK" Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.BaseNewSheet
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.BaseNewSheet
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''CHAVE PRIMARIA PRETA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                    pk = Cells(linha.Row, coluna.Column)
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE valor003 = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListBaseNewSheet
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditTbProdutos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Tb"
nomedasheet = "Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.TbProdutos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.TbProdutos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKProduto = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListTbProdutos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsPartes_Produtos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Partes_Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsPartes_Produtos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsPartes_Produtos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKParte_Produto = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsPartes_Produtos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsClientes()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Clientes"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        If ActiveSheet.optButton1CnsClientes = True Then
            SQL = "UPDATE " & tipotabela & nomedasheet & "PF SET "
        ElseIf ActiveSheet.optButton2CnsClientes = True Then
            SQL = "UPDATE " & tipotabela & nomedasheet & "PJ SET "
        End If
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsClientes
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsClientes
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKCliente = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsClientes
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsPedidos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Pedidos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Or Left(nomecoluna, 2) = "PK" Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsPedidos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsPedidos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''CHAVE PRIMARIA PRETA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(90, 90, 90) And Left(nomecoluna, 2) = "PK" Then
                    pk = Cells(linha.Row, coluna.Column)
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKPedido = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsPedidos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub


Sub EditCnsInsumos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Insumos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsInsumos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsInsumos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKInsumo = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsInsumos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsInsumos_Produtos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Insumos_Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsInsumos_Produtos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsInsumos_Produtos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKInsumo_Produto = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsInsumos_Produtos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsCortes_Insumos_Produtos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Cortes_Insumos_Produtos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsCortes_Insumos_Produtos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsCortes_Insumos_Produtos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKCorte_Insumo_Produto = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsCortes_Insumos_Produtos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsProducao_Pedidos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Producao_Pedidos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Or Left(nomecoluna, 2) = "PK" Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsProducao_Pedidos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsProducao_Pedidos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                
                ''CHAVE PRIMARIA PRETA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(90, 90, 90) And Left(nomecoluna, 2) = "PK" Then
                    pk = Cells(linha.Row, coluna.Column)
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKProducao_Pedido = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsProducao_Pedidos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsCompras_Insumos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Compras_Insumos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Or Left(nomecoluna, 2) = "PK" Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsCompras_Insumos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsCompras_Insumos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''CHAVE PRIMARIA PRETA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(90, 90, 90) And Left(nomecoluna, 2) = "PK" Then
                    pk = Cells(linha.Row, coluna.Column)
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKCompra_Insumo = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsCompras_Insumos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub EditCnsEstoque_Insumos()

Application.ScreenUpdating = False
Application.EnableEvents = False

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''DECLARA TODAS AS VARIÁVEIS GLOBAIS DO MÓDULO 'GENERAL'
''------------------------------------------------------
Call General.DeclarePublic

tipotabela = "Cns"
nomedasheet = "Estoque_Insumos"

'Seleciona a planilha a ser utilizada
'---------------------------------------------
ThisWorkbook.Worksheets(nomedasheet).Activate

''CRIA A CONEXÃO COM O BANCO DE DADOS
''-----------------------------------
Call General.DefineDBConection

''CONECTA AO BANCO DE DADOS
''-------------------------
Call General.ConectDB

ninputedit = 2 'Indica para a validação de dados que é um input de um 'Edit'
cor = "" 'Indica para a Validação de Dados a Cor da célula
arrignocel = Array()
arrignoval = Array()
editouqlqcoisa = False
msgignorados = "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf

'Define o comando SQL
'----------------------------------------------
For Each linha In Range(tipotabela & nomedasheet).Rows

    editou = False
    
    If ActiveSheet.Range("A" & linha.Row) = "MODIFICADO" And ActiveSheet.Range("A" & linha.Row).Interior.Color = RGB(255, 140, 50) Then
        
        SQL = "UPDATE " & tipotabela & nomedasheet & " SET "
        
        For Each coluna In Range(tipotabela & nomedasheet).Columns
        
            nomecoluna = Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column)
            
            If Cells(linha.Row, coluna.Column).Interior.Color = RGB(255, 140, 50) Or Left(Cells(Range("New" & tipotabela & nomedasheet).Row, coluna.Column), 2) = "PK" Then
            
                ''BRANCO
                If Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(240, 240, 240) Then
                    SQL = SQL & nomecoluna & " = '" & Cells(linha.Row, coluna.Column) & "', "
                    editou = True
                ''AMARELO
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 240, 205) Then
                    cor = RGB(255, 240, 205)
                    valoramarelo = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsEstoque_Insumos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valoramarelo & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valoramarelo
                        Else
                            SQL = SQL & nomecoluna & " = '" & valoramarelo & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula AMARELA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''LARANJA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(255, 230, 205) Then
                    cor = RGB(255, 230, 205)
                    valorlaranja = Cells(linha.Row, coluna.Column)
                    Call ValidateData.CnsEstoque_Insumos
                    If existe = True Then
                        If ignorado = True Then
                            msgignorados = msgignorados & Cells(linha.Row, coluna.Column).Address(0, 0) & "['" & valorlaranja & "'], "
                            ReDim Preserve arrignocel(-1 To UBound(arrignocel) + 1)
                            ReDim Preserve arrignoval(-1 To UBound(arrignoval) + 1)
                            arrignocel(UBound(arrignocel)) = Cells(linha.Row, coluna.Column).Address(0, 0)
                            arrignoval(UBound(arrignoval)) = valorlaranja
                        Else
                            SQL = SQL & nomecoluna & " = '" & valorlaranja & "', "
                            editou = True
                        End If
                    ElseIf existe = False Then
                        MsgBox ("Não existe um procedimento de validação de dados para a coluna '" & nomecoluna & "'!" & vbCrLf & "É necessário que toda célula LARANJA tenha uma validação de dados quando se modifica registro!")
                        cn.Close
                        End
                    End If
                    
                ''CHAVE PRIMARIA PRETA
                ElseIf Cells(Range("New" & tipotabela & nomedasheet).Row + 1, coluna.Column).Interior.Color = RGB(90, 90, 90) Then
                    pk = Cells(linha.Row, coluna.Column)
                    
                ''PRETO, VERDE e AZUL (ou qualquer outra Cor não especificada)
                Else
                    GoTo nextcoluna
                    
                End If
                
            End If
        
nextcoluna:
        Next
        
        If editou = True Then
            editouqlqcoisa = True
            SQL = Left(SQL, Len(SQL) - 2) & " WHERE PKInsumo = " & Cells(linha.Row, Range(tipotabela & nomedasheet).Column)
            
            ''EXECUTA O COMANDO SQL NO BANCO DE DADOS
            ''---------------------------------------
            Call General.OpenRS
        End If
        
        ''LIMPA A CÉLULA DA COLUNA 'A' E A LINHA MODIFICADA PARA INDICAR QUE OS DADOS JÁ FORAM MODIFICADOS
        ''-----------------------------------------------------------------------------------------------
        ActiveSheet.Range("A" & linha.Row).ClearContents
        ActiveSheet.Range("A" & linha.Row).ClearFormats
        Rows(ActiveSheet.Range("A" & linha.Row).Row).Interior.Color = xlNone
        
    End If

Next

''DESCONECTA DO BANCO DE DADOS
''----------------------------------
Call General.DisconectDB

''LISTA A TABELA ATUALIZADA NA PLANILHA SE ALGUM VALOR TIVER SIDO ALTERADO
''------------------------------------------------------------------------
If editouqlqcoisa = True Then
    Call ListData.ListCnsEstoque_Insumos
End If

'DESPROTEGE A PLANILHA
'---------------------
Call General.UnprotectSheet

''ALTERA OS VALORES DAS CÉLULAS IGNORADAS PARA O VALOR DIGITADO PELO USUÁRIO ANTERIORMENTE, E AS COLORE DE VERMELHO
''-----------------------------------------------------------------------------------------------------------------
Application.EnableEvents = False 'O ListData define como True, então é necessário desativar os eventos novamente
For vai = LBound(arrignocel) + 1 To UBound(arrignocel)
    Range(arrignocel(vai)).Interior.Color = RGB(255, 55, 40)
    Range(arrignocel(vai)).Value = arrignoval(vai)
Next

'PROTEGE A PLANILHA
'------------------
Call General.ProtectSheet

''EXIBE A MENSAGEM DAS CÉLULAS QUE FORAM IGNORADAS NA MODIFICAÇÃO
''---------------------------------------------------------------
If msgignorados <> "As seguintes células foram ignoradas pois não permitem os respectivos valores:" & vbCrLf Then
    msgignorados = Left(msgignorados, Len(msgignorados) - 2)
    MsgBox msgignorados
End If

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

