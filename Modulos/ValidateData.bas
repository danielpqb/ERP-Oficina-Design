Attribute VB_Name = "ValidateData"

Dim SQLAMARELO      As String
Dim SQLPRETO             As String
Dim SQLLARANJA         As String

Sub BaseNewSheet()

'valor002 = Tb
'valor003 = PKProduto
'BaseNewSheet = TbProdutos

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''valor003
        If nomecoluna = "valor003" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(valor003) FROM " & "valor002" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            pk = valorpreto
            rs.Close
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
    
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub TbProdutos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL

        On Error GoTo erro
    
        ''PKProduto
        If nomecoluna = "PKProduto" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKProduto) FROM " & tipotabela & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            rs.Close
            Exit Sub
        End If
        
        ''Cod_Produto
        If nomecoluna = "Cod_Produto" Then
            existe = True
            bloqueado = False
            ignorado = True
            Exit Sub
        End If
        
        ''DataRegistro_Produto
        If nomecoluna = "DataRegistro_Produto" Then
            valorpreto = Now()
            existe = True
            bloqueado = False
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
    
        ''Categoria_Produto
        If nomecoluna = "Categoria_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM Tb" & nomecoluna & " WHERE Abreviacao_Categoria_Produto='" & UCase(valoramarelo) & "';"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            ElseIf valoramarelo = "" Then
                bloqueado = False
                valoramarelo = "-"
            End If
            rs.Close
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
        ''Linha_Produto
        If nomecoluna = "Linha_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM Tb" & nomecoluna & " WHERE Abreviacao_Linha_Produto='" & UCase(valoramarelo) & "';"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            ElseIf valoramarelo = "" Then
                bloqueado = False
                valoramarelo = "-"
            End If
            rs.Close
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
        ''Tipo_Produto
        If nomecoluna = "Tipo_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM Tb" & nomecoluna & " WHERE Abreviacao_Tipo_Produto='" & UCase(valoramarelo) & "';"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            ElseIf valoramarelo = "" Then
                bloqueado = False
                valoramarelo = "-"
            End If
            rs.Close
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
        ''SR, MC, VD, ES
        If nomecoluna = "SR" Or nomecoluna = "MC" Or nomecoluna = "VD" Or nomecoluna = "ES" Then
            existe = True
            If valoramarelo <> "" Then
                valoramarelo = "X"
            End If
            bloqueado = False
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
        ''Categoria_Produto
        If nomecoluna = "Categoria_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM Tb" & nomecoluna & " WHERE Abreviacao_Categoria_Produto='" & UCase(valoramarelo) & "';"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            ElseIf valoramarelo = "" Then
                ignorado = False
                valoramarelo = "-"
            End If
            rs.Close
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
        ''Linha_Produto
        If nomecoluna = "Linha_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM Tb" & nomecoluna & " WHERE Abreviacao_Linha_Produto='" & UCase(valoramarelo) & "';"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            ElseIf valoramarelo = "" Then
                ignorado = False
                valoramarelo = "-"
            End If
            rs.Close
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
        ''Tipo_Produto
        If nomecoluna = "Tipo_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM Tb" & nomecoluna & " WHERE Abreviacao_Tipo_Produto='" & UCase(valoramarelo) & "';"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            ElseIf valoramarelo = "" Then
                ignorado = False
                valoramarelo = "-"
            End If
            rs.Close
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
        ''SR, MC, VD, ES
        If nomecoluna = "SR" Or nomecoluna = "MC" Or nomecoluna = "VD" Or nomecoluna = "ES" Then
            existe = True
            If valoramarelo <> "" Then
                valoramarelo = "X"
            End If
            ignorado = False
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsPartes_Produtos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKProduto
        If nomecoluna = "PKParte_Produto" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKParte_Produto) FROM " & tipotabela & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            rs.Close
            Exit Sub
        End If
        
        ''Cod_Produto
        If nomecoluna = "Cod_Produto" Then
            existe = True
            bloqueado = False
            ignorado = True
            Exit Sub
        End If
        
        ''Descricao_Produto
        If nomecoluna = "Descricao_Produto" Then
            existe = True
            bloqueado = False
            ignorado = True
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''FK_PRODUTO
        ''----------
        If nomecoluna = "FKProduto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbProdutos" & " WHERE PKProduto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''SR, MC, PT, VD, ES
        ''------------------
        If nomecoluna = "SR" Or nomecoluna = "MC" Or nomecoluna = "PT" Or nomecoluna = "VD" Or nomecoluna = "ES" Then
            existe = True
            bloqueado = False
            If valoramarelo <> "" Then
                valoramarelo = "X"
            End If
            Exit Sub
        End If
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        ''FK_PRODUTO
        ''----------
        If nomecoluna = "FKProduto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbProdutos" & " WHERE PKProduto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''SR, MC, PT, VD, ES
        ''------------------
        If nomecoluna = "SR" Or nomecoluna = "MC" Or nomecoluna = "PT" Or nomecoluna = "VD" Or nomecoluna = "ES" Then
            existe = True
            ignorado = False
            If valoramarelo <> "" Then
                valoramarelo = "X"
            End If
            Exit Sub
        End If

    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsClientes()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKCliente
        If nomecoluna = "PKCliente" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKCliente) FROM " & "Tb" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            rs.Close
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
    
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsPedidos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKPedido
        If nomecoluna = "PKPedido" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKPedido) FROM " & "Tb" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            pk = valorpreto
            rs.Close
            Exit Sub
        End If
        
        ''Numero_Pedido
        If nomecoluna = "Numero_Pedido" Then
            existe = True
            bloqueado = False
            If numpedido = 0 Then 'Se for a primeira vez que o programa entra nessa condição
                SQLPRETO = "SELECT MAX(Numero_Pedido) FROM " & "Tb" & nomedasheet
                rs.Open SQLPRETO, cn
                valorpreto = rs(0) + 1
                rs.Close
                numpedido = valorpreto 'Guarda o numero do pedido caso exista outra passada nessa condição
            Else
                valorpreto = numpedido
            End If
            Exit Sub
        End If
        
        ''Nome_Cliente, Cod_Produto, Descricao_Produto
        If nomecoluna = "Nome_Cliente" Or nomecoluna = "Cod_Produto" Or nomecoluna = "Descricao_Produto" Then
            existe = True
            bloqueado = False
            ignorado = True
            valorpreto = ""
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''FKCliente
        If nomecoluna = "FKCliente" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbClientes" & " WHERE PKCliente=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If

        ''FKProduto
        If nomecoluna = "FKProduto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbProdutos" & " WHERE PKProduto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Quantidade
        If nomecoluna = "Quantidade" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                valoramarelo = CInt(valoramarelo)
                If valoramarelo < 1 Then valoramarelo = 1
            Else
                valoramarelo = 1
            End If
            bloqueado = False
            Exit Sub
        End If
        
        ''Data_Venda
        If nomecoluna = "Data_Venda" Then
            existe = True
            If IsDate(valorlaranja) = False Then
                valorlaranja = Now()
            End If
            bloqueado = False
            Exit Sub
        End If
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
        ''FKCliente
        If nomecoluna = "FKCliente" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbClientes" & " WHERE PKCliente=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            End If
            rs.Close
            Exit Sub
        End If

        ''FKProduto
        If nomecoluna = "FKProduto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbProdutos" & " WHERE PKProduto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Quantidade
        If nomecoluna = "Quantidade" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                valoramarelo = CInt(valoramarelo)
                If valoramarelo < 1 Then valoramarelo = 1
            Else
                valoramarelo = 1
            End If
            ignorado = False
            Exit Sub
        End If
        
        ''Data_Venda
        If nomecoluna = "Data_Venda" Then
            existe = True
            If IsDate(valorlaranja) = False And valorlaranja <> "" Then
                valorlaranja = Now()
            End If
            ignorado = False
            Exit Sub
        End If
        
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''QntProducao
        If nomecoluna = "QntProducao" Then
            existe = True
            valorlaranja = 0
            bloqueado = False
            Exit Sub
        End If
        
        ''Cancelado
        If nomecoluna = "Cancelado" Then
            existe = True
            valorlaranja = ""
            bloqueado = False
            Exit Sub
        End If
        
        ''Data_Entrega
        If nomecoluna = "Data_Entrega" Then
            existe = True
            valorlaranja = ""
            bloqueado = False
            Exit Sub
        End If
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        ''QntProducao
        If nomecoluna = "QntProducao" Then
            existe = True
            SQLLARANJA = "SELECT COUNT(*) FROM TbProducao_Pedidos WHERE FKPedido=" & pk & ";"
            rs.Open SQLLARANJA, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                emprod = rs.Fields(0) 'Quantos produtos ja estao em producao
            End If
            rs.Close
            '
            SQLLARANJA = "SELECT Quantidade FROM TbPedidos WHERE PKPedido=" & pk & ";"
            rs.Open SQLLARANJA, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                qnt = rs.Fields(0) 'Quantos produtos devem ser produzidos para esse pedido
            End If
            rs.Close
            '
            If IsNumeric(valorlaranja) = True Then
                valorlaranja = CInt(valorlaranja)
                If valorlaranja < 0 Then valorlaranja = 0
                qntintencao = valorlaranja 'Quantidade de Produtos que o usuario deseja adicionar na producao
                
                ''Retorna o valor da QntProducao antes de modificar
                SQLLARANJA = "SELECT QntProducao FROM " & "TbPedidos" & " WHERE PKPedido=" & pk & ";"
                rs.Open SQLLARANJA, cn
                qntanterior = CInt(rs(0))
                rs.Close
                
                ''Proximo PKProducao_Pedido
                SQLLARANJA = "SELECT MAX(PKProducao_Pedido) FROM TbProducao_Pedidos"
                rs.Open SQLLARANJA, cn
                proxpk = rs(0) + 1
                rs.Close
                ''Proximo Num_Producao
                SQLLARANJA = "SELECT MAX(Num_Producao) FROM TbProducao_Pedidos WHERE FKPedido=" & pk & ";"
                rs.Open SQLLARANJA, cn
                If IsNumeric(rs(0)) = True Then 'Se o valor procurado for encontrado, então...
                    proxnumprod = rs(0) + 1
                Else
                    proxnumprod = 1
                End If
                rs.Close
                ''Qual PKProduto estamos trabalhando
                SQLLARANJA = "SELECT FKProduto FROM TbPedidos WHERE PKPedido=" & pk & ";"
                rs.Open SQLLARANJA, cn
                pkprod = rs(0)
                rs.Close
                ''Se tem SR, MC, VD, ES
                SQLLARANJA = "SELECT SR, MC, VD, ES FROM TbProdutos WHERE PKProduto=" & pkprod & ";"
                rs.Open SQLLARANJA, cn
                If rs(0) = "X" Then
                    sr = ""
                Else
                    sr = "-"
                End If
                If rs(1) = "X" Then
                    mc = ""
                Else
                    mc = "-"
                End If
                If rs(2) = "X" Then
                    vd = ""
                Else
                    vd = "-"
                End If
                If rs(3) = "X" Then
                    es = ""
                Else
                    es = "-"
                End If
                rs.Close
                
                If valorlaranja + emprod <= qnt And valorlaranja > 0 Then
                    For i = 1 To valorlaranja
                        SQLLARANJA = "INSERT INTO TbProducao_Pedidos (PKProducao_Pedido, FKPedido, Num_Producao, Inicio_SR, Fim_SR, Inicio_MC, Fim_MC, Inicio_PT_SR, " & _
                                                                                                                            "Fim_PT_SR, Inicio_PT_MC, Fim_PT_MC, Inicio_VD, Fim_VD, Inicio_ES, Fim_ES)" & _
                                                        "VALUES ('" & proxpk & "', '" & pk & "', '" & proxnumprod & "', '" & sr & "', '" & sr & "', '" & mc & "', '" & mc & "', '" & sr & _
                                                                            "', '" & sr & "', '" & mc & "', '" & mc & "', '" & vd & "', '" & vd & "', '" & es & "', '" & es & "');"
                        rs.Open SQLLARANJA, cn
                        proxpk = proxpk + 1
                        proxnumprod = proxnumprod + 1
                    Next
                    valorlaranja = valorlaranja + qntanterior
                    ignorado = False
                ElseIf valorlaranja + emprod > qnt And valorlaranja > 0 Then
                    For i = emprod To qnt - 1
                        SQLLARANJA = "INSERT INTO TbProducao_Pedidos (PKProducao_Pedido, FKPedido, Num_Producao, Inicio_SR, Fim_SR, Inicio_MC, Fim_MC, Inicio_PT_SR, " & _
                                                                                                                            "Fim_PT_SR, Inicio_PT_MC, Fim_PT_MC, Inicio_VD, Fim_VD, Inicio_ES, Fim_ES)" & _
                                                        "VALUES ('" & proxpk & "', '" & pk & "', '" & proxnumprod & "', '" & sr & "', '" & sr & "', '" & mc & "', '" & mc & "', '" & sr & _
                                                                            "', '" & sr & "', '" & mc & "', '" & mc & "', '" & vd & "', '" & vd & "', '" & es & "', '" & es & "');"
                        rs.Open SQLLARANJA, cn
                        proxpk = proxpk + 1
                        proxnumprod = proxnumprod + 1
                    Next
                    valorlaranja = qnt
                    ignorado = False
                End If
            End If
            Exit Sub
        End If
        
        ''Cancelado
        If nomecoluna = "Cancelado" Then
            existe = True
            If IsDate(valorlaranja) = False And valorlaranja <> "" Then
                valorlaranja = Now()
            End If
            ignorado = False
            Exit Sub
        End If
        
        ''Data_Entrega
        If nomecoluna = "Data_Entrega" Then
            existe = True
            If IsDate(valorlaranja) = False And valorlaranja <> "" Then
                valorlaranja = Now()
            End If
            ignorado = False
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''Data_LimiteEntrega
        If nomecoluna = "Data_LimiteEntrega" Then
            existe = True
            ignorado = False
            bloqueado = False
            If IsDate(valorverde) = False Then
                valorverde = Date + 30
            End If
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub


Sub CnsInsumos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKInsumo
        If nomecoluna = "PKInsumo" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKInsumo) FROM " & "Tb" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            rs.Close
            Exit Sub
        End If
        
        ''Preco_UnMed
        If nomecoluna = "Preco_UnMed" Then
            existe = True
            bloqueado = False
            ignorado = True
            valorpreto = ""
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
    
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsInsumos_Produtos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKInsumo_Produto
        If nomecoluna = "PKInsumo_Produto" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKInsumo_Produto) FROM " & "Cns" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            rs.Close
            Exit Sub
        End If
        
        ''Descricao_Insumo, Cod_Produto, Descricao_Produto, UnMed_Insumo, Custo
        If nomecoluna = "Descricao_Insumo" Or nomecoluna = "Cod_Produto" Or nomecoluna = "Descricao_Produto" Or nomecoluna = "UnMed_Insumo" Or nomecoluna = "Custo" Then
            existe = True
            bloqueado = False
            ignorado = True
            Exit Sub
        End If
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
    
        ''FKProduto
        If nomecoluna = "FKProduto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbProdutos" & " WHERE PKProduto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''FKInsumo
        If nomecoluna = "FKInsumo" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbInsumos" & " WHERE PKInsumo=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Qnt
        If nomecoluna = "Qnt" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                bloqueado = False
            End If
            Exit Sub
        End If
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
    ''FKProduto
        If nomecoluna = "FKProduto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbProdutos" & " WHERE PKProduto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''FKInsumo
        If nomecoluna = "FKInsumo" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbInsumos" & " WHERE PKInsumo=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Qnt
        If nomecoluna = "Qnt" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                ignorado = False
            End If
            Exit Sub
        End If
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsCortes_Insumos_Produtos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKCorte_Insumo_Produto
        If nomecoluna = "PKCorte_Insumo_Produto" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKCorte_Insumo_Produto) FROM " & "Cns" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            rs.Close
            Exit Sub
        End If
        
        ''Descricao_Produto, Descricao_Insumo, Cod_Produto
        If nomecoluna = "Descricao_Produto" Or nomecoluna = "Descricao_Insumo" Or nomecoluna = "Cod_Produto" Then
            existe = True
            bloqueado = False
            ignorado = True
            valorpreto = ""
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
        
        ''FKInsumo_Produto
        If nomecoluna = "FKInsumo_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbInsumos_Produtos" & " WHERE PKInsumo_Produto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Qnt
        If nomecoluna = "Qnt" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                bloqueado = False
            End If
            Exit Sub
        End If
    
        On Error GoTo erro
    
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
        
        ''FKInsumo_Produto
        If nomecoluna = "FKInsumo_Produto" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbInsumos_Produtos" & " WHERE PKInsumo_Produto=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                ignorado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Qnt
        If nomecoluna = "Qnt" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                ignorado = False
            End If
            Exit Sub
        End If
    
        On Error GoTo ignorar
    
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsProducao_Pedidos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
    
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
        
        On Error GoTo ignorar
        
        ''Material
        If nomecoluna = "Material" Then
            existe = True
            If valoramarelo <> "" Then
                valoramarelo = "X"
            End If
            ignorado = False
            valoramarelo = UCase(valoramarelo)
            
            SQLAMARELO = "SELECT MAX(PKLogInOutEstoque) FROM TbLogInOutEstoque"
            rs.Open SQLAMARELO, cn
            proxpk = rs(0) + 1 'Proximo PKLogInOutEstoque
            rs.Close
            
            SQLAMARELO = "SELECT Material FROM TbProducao_Pedidos WHERE PKProducao_Pedido = " & pk & ";"
            rs.Open SQLAMARELO, cn
            Material = rs(0)
            rs.Close

            SQLMITO = "SELECT DISTINCTROW TbInsumos.PKInsumo, Round((([TbInsumos_Produtos].[Qnt])/[TbInsumos].[UnMedPorCompra])/(1-[TbInsumos].[Perda_Insumo]),4) AS Qnt, TbProducao_Pedidos.PKProducao_Pedido FROM (TbProdutos INNER JOIN (TbInsumos INNER JOIN TbInsumos_Produtos ON TbInsumos.PKInsumo = TbInsumos_Produtos.FKInsumo) ON TbProdutos.PKProduto = TbInsumos_Produtos.FKProduto) INNER JOIN (TbPedidos INNER JOIN TbProducao_Pedidos ON TbPedidos.PKPedido = TbProducao_Pedidos.FKPedido) ON TbProdutos.PKProduto = TbPedidos.FKProduto GROUP BY TbInsumos.PKInsumo, Round((([TbInsumos_Produtos].[Qnt])/[TbInsumos].[UnMedPorCompra])/(1-[TbInsumos].[Perda_Insumo]),4), TbProducao_Pedidos.PKProducao_Pedido HAVING (((TbProducao_Pedidos.PKProducao_Pedido)=" & pk & "));"
            rs.Open SQLMITO, cn
            Do While rs.EOF = False
                If Material = "" Or IsNull(Material) = True Then
                    SQLAMARELO = "INSERT INTO TbLogInOutEstoque (PKLogInOutEstoque, FKInsumo, Qnt, Data_LogInOut, Descricao_LogInOut, FKProducao_Pedido)" & _
                                                            "VALUES ('" & proxpk & "', '" & rs(0) & "', '" & -rs(1) & "', '" & Now() & "', '" & "Material comprometido" & "', '" & pk & "');"
                    rs2.Open SQLAMARELO, cn
                    proxpk = proxpk + 1
                    rs.MoveNext
                Else
                    SQLAMARELO = "INSERT INTO TbLogInOutEstoque (PKLogInOutEstoque, FKInsumo, Qnt, Data_LogInOut, Descricao_LogInOut, FKProducao_Pedido)" & _
                                                            "VALUES ('" & proxpk & "', '" & rs(0) & "', '" & rs(1) & "', '" & Now() & "', '" & "Material descomprometido" & "', '" & pk & "');"
                    rs2.Open SQLAMARELO, cn
                    proxpk = proxpk + 1
                    rs.MoveNext
                End If
            Loop
            rs.Close
            Exit Sub
        End If
        ''----
        
        ''Envio
        If nomecoluna = "Envio" Then
            existe = True
            ignorado = False
            If IsDate(valoramarelo) = False And valoramarelo <> "" Then
                valoramarelo = Now()
            End If
            
            SQLAMARELO = "SELECT MAX(PKLogInOutEstoque) FROM TbLogInOutEstoque"
            rs.Open SQLAMARELO, cn
            proxpk = rs(0) + 1 'Proximo PKLogInOutEstoque
            rs.Close
            
            SQLAMARELO = "SELECT Material FROM TbProducao_Pedidos WHERE PKProducao_Pedido = " & pk & ";"
            rs.Open SQLAMARELO, cn
            Material = rs(0)
            rs.Close

            SQLMITO = "SELECT DISTINCTROW TbInsumos.PKInsumo, Round((([TbPedidos].[Quantidade]*[TbInsumos_Produtos].[Qnt])/[TbInsumos].[UnMedPorCompra])/(1-[TbInsumos].[Perda_Insumo]),4) AS Qnt, TbProducao_Pedidos.PKProducao_Pedido FROM TbCompras_Insumos, (TbProdutos INNER JOIN (TbInsumos INNER JOIN TbInsumos_Produtos ON TbInsumos.PKInsumo = TbInsumos_Produtos.FKInsumo) ON TbProdutos.PKProduto = TbInsumos_Produtos.FKProduto) INNER JOIN (TbPedidos INNER JOIN TbProducao_Pedidos ON TbPedidos.PKPedido = TbProducao_Pedidos.FKPedido) ON TbProdutos.PKProduto = TbPedidos.FKProduto GROUP BY TbInsumos.PKInsumo, Round((([TbPedidos].[Quantidade]*[TbInsumos_Produtos].[Qnt])/[TbInsumos].[UnMedPorCompra])/(1-[TbInsumos].[Perda_Insumo]),4), TbProducao_Pedidos.PKProducao_Pedido HAVING (((TbProducao_Pedidos.PKProducao_Pedido)=" & pk & "));"
            rs.Open SQLMITO, cn
            If Material = "" Or IsNull(Material) = True Then
                SQLAMARELO = "UPDATE TbProducao_Pedidos SET Material = 'X' WHERE PKProducao_Pedido=" & pk & ";"
                rs2.Open SQLAMARELO, cn
                Do While rs.EOF = False
                    SQLAMARELO = "INSERT INTO TbLogInOutEstoque (PKLogInOutEstoque, FKInsumo, Qnt, Data_LogInOut, Descricao_LogInOut, FKProducao_Pedido)" & _
                                                            "VALUES ('" & proxpk & "', '" & rs(0) & "', '" & -rs(1) & "', '" & Now() & "', '" & "Material comprometido após envio" & "', '" & pk & "');"
                    rs2.Open SQLAMARELO, cn
                    proxpk = proxpk + 1
                    rs.MoveNext
                Loop
            End If
            rs.Close
            Exit Sub
        End If
        ''----
        
        ''Plano_Corte
        If nomecoluna = "Plano_Corte" Then
            existe = True
            If valoramarelo <> "" Then
                valoramarelo = "X"
            End If
            ignorado = False
            valoramarelo = UCase(valoramarelo)
            Exit Sub
        End If
        ''----
        
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsCompras_Insumos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''PKCompra_Insumo
        If nomecoluna = "PKCompra_Insumo" Then
            existe = True
            bloqueado = False
            SQLPRETO = "SELECT MAX(PKCompra_Insumo) FROM " & "Cns" & nomedasheet
            rs.Open SQLPRETO, cn
            valorpreto = rs(0) + 1
            pk = valorpreto
            rs.Close
            Exit Sub
        End If
        
        ''Categoria_Insumo, Descricao_Insumo, UnCompra_Insumo, Nome_Fornecedor
        If nomecoluna = "Categoria_Insumo" Or nomecoluna = "Descricao_Insumo" Or nomecoluna = "UnCompra_Insumo" Or nomecoluna = "Nome_Fornecedor" Then
            existe = True
            bloqueado = False
            ignorado = True
            valorpreto = ""
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        ''FKInsumo
        If nomecoluna = "FKInsumo" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbInsumos" & " WHERE PKInsumo=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''FKFornecedor
        If nomecoluna = "FKFornecedor" Then
            existe = True
            SQLAMARELO = "SELECT * FROM " & "TbFornecedores" & " WHERE PKFornecedor=" & valoramarelo & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                bloqueado = False
            End If
            rs.Close
            Exit Sub
        End If
        
        ''Qnt
        If nomecoluna = "Qnt" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                valoramarelo = CInt(valoramarelo)
            Else
                valoramarelo = 0
            End If
            bloqueado = False
            Exit Sub
        End If
        
        ''Preco
        If nomecoluna = "Preco" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
            Else
                valoramarelo = 0
            End If
            bloqueado = False
            Exit Sub
        End If
    
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
    
        ''FKInsumo
        If nomecoluna = "FKInsumo" Then
            existe = True
            SQLAMARELO = "SELECT Recebido FROM " & "TbCompras_Insumos" & " WHERE PKCompra_Insumo=" & pk & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                If IsNull(rs(0)) Or rs(0) = "" Then
                    rs.Close
                    SQLAMARELO = "SELECT * FROM " & "TbInsumos" & " WHERE PKInsumo=" & valoramarelo & ";"
                    rs.Open SQLAMARELO, cn
                    If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                        ignorado = False
                    End If
                    rs.Close
                Else
                    rs.Close
                End If
            Else
                rs.Close
            End If
            Exit Sub
        End If
        
        ''FKFornecedor
        If nomecoluna = "FKFornecedor" Then
            existe = True
            SQLAMARELO = "SELECT Recebido FROM " & "TbCompras_Insumos" & " WHERE PKCompra_Insumo=" & pk & ";"
            rs.Open SQLAMARELO, cn
            If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                If IsNull(rs(0)) Or rs(0) = "" Then
                    rs.Close
                    SQLAMARELO = "SELECT * FROM " & "TbFornecedores" & " WHERE PKFornecedor=" & valoramarelo & ";"
                    rs.Open SQLAMARELO, cn
                    If rs.EOF = False Then 'Se o valor procurado for encontrado, então...
                        ignorado = False
                    End If
                    rs.Close
                Else
                    rs.Close
                End If
            Else
                rs.Close
            End If
            Exit Sub
        End If
        
        ''Qnt
        If nomecoluna = "Qnt" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
                valoramarelo = CInt(valoramarelo)
            Else
                valoramarelo = 0
            End If
            ignorado = False
            Exit Sub
        End If
        
        ''Preco
        If nomecoluna = "Preco" Then
            existe = True
            If IsNumeric(valoramarelo) = True Then
            Else
                valoramarelo = 0
            End If
            ignorado = False
            Exit Sub
        End If
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = True    'Define se o dado será excluído do comando SQL
        existe = True
        bloqueado = False
    
        On Error GoTo erro
        
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        ''Recebido
        If nomecoluna = "Recebido" Then
            existe = True
            If valorlaranja <> "" Then
                valorlaranja = "X"
            End If
            ignorado = False
            valorlaranja = UCase(valorlaranja)
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub

Sub CnsEstoque_Insumos()

'-----------------------------------------------------------------------------------------
'PROCEDIMENTOS NECESSÁRIOS PARA VERIFICAR E VALIDAR OS DADOS NO 'EDITDATA' E NO 'NEWINPUT'
'-----------------------------------------------------------------------------------------

existe = False      'Define se existe um procedimento de validação de dados
ignorado = True     'Define se o dado será excluído do comando SQL
bloqueado = True    'Define se o programa será parado

''--------------------''
''        PRETO       ''
''--------------------''
If cor = RGB(90, 90, 90) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (PRETO) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub

''--------------------''
''       AMARELO      ''
''--------------------''
ElseIf cor = RGB(255, 240, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
    
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        
    Else
    MsgBox "Ocorreu um erro ao definir se a Validação de dados (AMARELO) era de um Novo Registro ou de um Edit!"
    End
        
    End If
    
    Exit Sub
        
''--------------------''
''       LARANJA      ''
''--------------------''
ElseIf cor = RGB(255, 230, 205) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
        
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
    ''-------------------''
    ''       EDIT        ''
    ''-------------------''
    ElseIf ninputedit = 2 Then
    
        bloqueado = False    'Define se o programa será parado
    
        On Error GoTo ignorar
        
        ''Contagem
        If nomecoluna = "Contagem" Then
            existe = True
            If IsNumeric(valorlaranja) = True Then
                ignorado = False
                SQLLARANJA = "UPDATE TbInsumos SET Data_Contagem = '" & Now() & "' WHERE PKInsumo=" & pk & ";"
                rs.Open SQLLARANJA, cn
            End If
            Exit Sub
        End If
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (LARANJA) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
''--------------------''
''        VERDE       ''
''--------------------''
ElseIf cor = RGB(225, 240, 220) Then

    ''-------------------''
    ''   NOVO REGISTRO   ''
    ''-------------------''
    If ninputedit = 1 Then
    
        ignorado = False    'Define se o dado será excluído do comando SQL
    
        On Error GoTo erro
        
        
        
    Else
        MsgBox "Ocorreu um erro ao definir se a Validação de dados (VERDE) era de um Novo Registro ou de um Edit!"
        End
    End If

    Exit Sub
    
Else

    MsgBox ("A cor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' não corresponde com nenhum procedimento de validação!")
    End

End If

Exit Sub 'Só coloquei de novo por segurança, mas não tem necessidade de ter mais um 'Exit Sub'

ignorar:
ignorado = True 'Não tem necessidade, mas coloquei por segurança
Exit Sub

erro:
MsgBox ("O valor da célula '" & Cells(ln, coluna.Column).Address(0, 0) & "' falhou ao ser definido!")
End

End Sub
