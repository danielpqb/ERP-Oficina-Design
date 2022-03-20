VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btExcluirCnsInsumos_Click()

Call DeleteData.DeleteDataCnsInsumos

End Sub

Private Sub btGravarCnsInsumos_Click()

Call NewInput.InputCnsInsumos

End Sub

Private Sub btListarCnsInsumos_Click()

Call ListData.ListCnsInsumos

End Sub

Private Sub btModificarCnsInsumos_Click()

Call EditData.EditCnsInsumos

End Sub

Private Sub txtboxQntNewRows_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then 'Se apertar a tecla 'Enter'

    Range("NewCnsInsumos").Select

End If

End Sub

Private Sub txtboxQntNewRows_LostFocus()

Application.EnableEvents = False

Call General.UnprotectSheet

''AJUSTA O NÚMERO DE LINHAS DE INPUT
''----------------------------------
If txtboxQntNewRows.Value < 1 Or txtboxQntNewRows.Value > 99 Then
    txtboxQntNewRows.Value = 1
End If

lcnd = Range("NewCnsInsumos").Row 'Linha do cabeçalho da inserção de novos dados
plt = Range("CnsInsumos").Row 'Primeira Linha da Tabela

nlind = plt - lcnd - 2 'Número de linhas para inserção de novos dados
nlfae = txtboxQntNewRows.Value - nlind 'Número de linhas que falta adicionar ou excluir

'Insere ou exclui o número de linhas necessário
If nlfae > 0 Then
    Rows(lcnd + 1 & ":" & lcnd + 1).Copy
        Rows(plt - 1 & ":" & plt - 1 + nlfae - 1).Insert Shift:=xlDown
        Rows(plt - 1 & ":" & plt - 1 + nlfae - 1).ClearContents
ElseIf nlfae < 0 Then
    Rows(plt - 2 + nlfae + 1 & ":" & plt - 2).Delete

End If

Call General.ProtectSheet

Application.EnableEvents = True

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

Application.EnableEvents = False

Call General.UnprotectSheet

''COLORIR AS CÉLULAS MODIFICADAS DA TABELA
''----------------------------------------
If Not Intersect(Target, Range("CnsInsumos")) Is Nothing Then 'Se a célula modificada estiver dentro da tabela, então entra na condição.

    If Not Intersect(Target, Rows(Range("CnsInsumos").Row)) Is Nothing Then 'Se a célula modificada for do cabeçalho da tabela, então entra na condição.
        MsgBox "Não é permitido modificar o cabeçalho da tabela!" & vbCrLf & "Os dados serão listados novamente para evitar bugs e crashes!"
        Call ListData.ListCnsInsumos
        GoTo fim
    End If
    If Not Intersect(Target, Columns(Range("CnsInsumos").Column)) Is Nothing Then 'Se a célula modificada estiver na primeira coluna da tabela, então entra na condição.
        MsgBox "Não é permitido modificar a primeira coluna da tabela!" & vbCrLf & "Os dados serão listados novamente para evitar bugs e crashes!"
        Call ListData.ListCnsInsumos
        GoTo fim
    End If

    Target.Interior.Color = RGB(255, 140, 50)
    For Each linha In Target.Rows
        With Range("A" & linha.Row)
            .Value = "MODIFICADO"
            .Font.Color = RGB(255, 140, 50)
            .Interior.Color = RGB(255, 140, 50)
            .Borders.LineStyle = xlContinuous
        End With
    Next

End If

If Not Intersect(Target, Rows(Range("NewCnsInsumos").Row & ":" & Range("NewCnsInsumos").Row + ActiveSheet.txtboxQntNewRows.Value)) Is Nothing Then 'Se a célula modificada estiver dentro da tabela de inserção de novos dados, então entra na condição.

    If Not Intersect(Target, Rows(Range("NewCnsInsumos").Row)) Is Nothing Then 'Se a célula modificada for do cabeçalho da tabela de inserção de dados, então entra na condição.
        MsgBox "Não é permitido modificar o cabeçalho da inserção de dados!" & vbCrLf & "Os dados serão listados novamente para evitar bugs e crashes!"
        Call ListData.ListCnsInsumos
        GoTo fim
    End If

    For Each linha In Target.Rows
        With Range("A" & linha.Row)
            .Value = "MODIFICADO"
            .Font.Color = RGB(255, 140, 50)
            .Interior.Color = RGB(255, 140, 50)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlRight
        End With
    Next

End If

Call General.ProtectSheet

fim:
Application.EnableEvents = True

End Sub
