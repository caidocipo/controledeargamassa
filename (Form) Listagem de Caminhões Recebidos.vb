Option Explicit

    Private arrData As Variant
    Private i As Long
    Private ws As Worksheet
    Private ws2 As Worksheet
    Private tbl As ListObject
    Private tblc As ListObject
    Private rngRecebID As Range
    Private rngData As Range
    Private rngConsID As Range

Private Sub UserForm_Activate()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If
    
    Set ws = ThisWorkbook.Worksheets("Recebimento")
    Set ws2 = ThisWorkbook.Worksheets("Consumo")
    Set tbl = ws.ListObjects("recebArgamassa")
    Set tblc = ws2.ListObjects("consArgamassa")
    Set rngRecebID = tbl.ListColumns("recebID").DataBodyRange
    Set rngData = tbl.ListColumns("Data").DataBodyRange
    Set rngConsID = tblc.ListColumns("ConsID").DataBodyRange

    arrData = rngRecebID.value
    
    ListBox1.Clear
    
    ' Adicionar os valores na ListBox1 em ordem decrescente
    For i = UBound(arrData, 1) To LBound(arrData, 1) Step -1
        ListBox1.AddItem arrData(i, 1)
    Next i
End Sub
'Funcionalidade que permite limpa o campo de pesquisa de data para permitir o preenchimento do usuário
Private Sub TextBox1_Enter()
    TextBox1.value = "" ' Remove o valor existente ao selecionar a TextBox
End Sub

'Funcionalidade que permite que os campos de data preencham a data junto ao usuário
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Verifica se o código ASCII do caractere digitado não é um número (exceto o backspace)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then ' Código ASCII para o backspace
            KeyAscii = 0 ' Cancela o caractere digitado
        End If
    End If
End Sub

Private Sub TextBox1_Change()
    Dim value As String
    Dim formattedValue As String
    
    ' Obtém o valor atual da TextBox
    value = TextBox1.value
    
    ' Remove todos os caracteres não numéricos
    value = Application.WorksheetFunction.Substitute(value, "/", "")
    
    ' Verifica se o valor é um número
    If IsNumeric(value) Then
        ' Adiciona automaticamente o caractere "/" ao digitar
        If Len(value) > 2 Then
            value = Left(value, 2) & "/" & Mid(value, 3)
        End If
        If Len(value) > 5 Then
            value = Left(value, 5) & "/" & Mid(value, 6)
        End If
    End If
    
    ' Atualiza o valor da TextBox
    TextBox1.value = value
End Sub

Private Sub ListBox1_Click()
    ' Verificar se um valor foi selecionado na ListBox1
    If ListBox1.ListIndex <> -1 Then
        ' Obter o valor selecionado na ListBox1
        Dim selectedValue As String
        selectedValue = ListBox1.value
        
        ' Procurar o valor selecionado na coluna "RecebID" da tabela "recebArgamassa"
        Dim rngFound As Range
        
        ' Procurar o valor selecionado na coluna "RecebID"
        Set rngFound = rngRecebID.Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Verificar se o valor foi encontrado
        If Not rngFound Is Nothing Then
            ' Atualizar os campos TextBox10 a TextBox16 com os valores correspondentes
            TextBox10.value = Format(rngFound.Offset(0, 2).value, "DD/MM/YYYY") ' Coluna "Data" no formato DD/MM/YYYY
            TextBox11.value = rngFound.Offset(0, 7).value ' Coluna "Caminhão"
            TextBox12.value = rngFound.Offset(0, 8).value ' Coluna "Pedido"
            TextBox13.value = rngFound.Offset(0, 9).value ' Coluna "Medição"
            TextBox14.value = rngFound.Offset(0, 10).value ' Coluna "Recebido (m³)"
            TextBox15.value = rngFound.Offset(0, 11).value ' Coluna "Consumido (m³)"
            TextBox16.value = rngFound.Offset(0, 12).value ' Coluna "Medido (m³)"
        End If
    End If
End Sub

Private Sub ListBox2_Click()
    ' Verificar se um valor foi selecionado na ListBox2
    If ListBox2.ListIndex <> -1 Then
        ' Obter o valor selecionado na ListBox2
        Dim selectedValue As String
        selectedValue = ListBox2.value
        
        ' Procurar o valor selecionado na coluna "RecebID" da tabela "recebArgamassa"
        Dim rngFound As Range
        
        ' Procurar o valor selecionado na coluna "RecebID"
        Set rngFound = rngConsID.Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Verificar se o valor foi encontrado
        If Not rngFound Is Nothing Then
            ' Atualizar os campos TextBox10 a TextBox16 com os valores correspondentes
            TextBox17.value = rngFound.Offset(0, 4).value ' Serviço
            TextBox18.value = rngFound.Offset(0, 6).value ' Bloco
            TextBox19.value = rngFound.Offset(0, 8).value ' Unidade
            TextBox20.value = rngFound.Offset(0, 9).value ' Consumido (m³)
            TextBox21.value = rngFound.Offset(0, 10).value ' Projetado (m³)
        End If
    End If
End Sub

Private Sub ListBox1_Change()
    Dim lastRow As Long
    Dim i As Long
    Dim searchValue As String
    Dim result As String
    
    ListBox2.Clear
    
    searchValue = ListBox1.value
    
    lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws2.Cells(i, "A").value = searchValue Then
            result = ws2.Cells(i, "B").value
            ListBox2.AddItem result
        End If
    Next i
End Sub

Private Sub CommandButton1_Click()
    Unload Me
    CaminhaoCadastro.Show vbModal
End Sub

Private Sub CommandButton2_Click()
    ' Verifica se um valor foi selecionado na ListBox1
    If ListBox1.ListIndex <> -1 Then
        ' Verifica se TextBox14 e TextBox15 contêm valores maiores que zero
        If IsNumeric(TextBox14.value) And IsNumeric(TextBox15.value) Then
            If CDbl(TextBox14.value) - CDbl(TextBox15.value) > 0 Then
                ' Carrega o novo formulário e transfere o valor selecionado para TextBox11
                ConsumoCadastro.TextBox1.value = ListBox1.value
                ConsumoCadastro.TextBox3.value = TextBox14.value
                ConsumoCadastro.TextBox4.value = TextBox15.value
                Unload Me
                ConsumoCadastro.Show vbModal
            Else
                MsgBox "Não há consumo disponível no caminhão selecionado.", vbExclamation, "Erro"
            End If
        Else
            MsgBox "Há um erro nos valores recebidos e consumidos do caminhão.", vbExclamation, "Erro"
        End If
    Else
        MsgBox "Selecione um caminhão para consumir.", vbExclamation, "Erro"
    End If
End Sub

Private Sub CommandButton3_Click()
    ' Limpar a ListBox1
    ListBox1.Clear
    
    ' Obter o valor digitado na TextBox1
    Dim searchValue As String
    searchValue = TextBox1.value
    
    ' Verificar se o valor digitado é uma data válida
    Dim searchDate As Date
    If IsDate(searchValue) Then
        searchDate = CDate(searchValue)
    Else
        searchDate = DateValue("01/01/1900") ' Definir uma data inválida
    End If
    
    ' Verificar se o valor digitado é vazio
    If Len(searchValue) = 0 Then
        ' Preencher a ListBox1 com todos os valores da coluna "recebID" em ordem decrescente
        Dim arrData2 As Variant
        Dim  As Long
        
        ' Inserir os valores na matriz
        arrData2 = rngRecebID.value
        
        ' Limpar a ListBox1
        ListBox1.Clear
        ' Adicionar os valores na ListBox1 em ordem decrescente
        For i = UBound(arrData2, 1) To LBound(arrData2, 1) Step -1
            ListBox1.AddItem arrData2(i, 1)
        Next i
    Else
        ' Verificar cada célula na coluna "Data" e adicionar o valor correspondente de "RecebID" na ListBox1
        Dim i As Long
        Dim valueArray() As Variant
        Dim index As Long
        
        ReDim valueArray(1 To rngRecebID.Cells.Count)
        index = 1
        
        For i = 1 To rngData.Cells.Count
            If rngData.Cells(i).value = searchDate Then
                valueArray(index) = rngRecebID.Cells(i).value
                index = index + 1
            End If
        Next i
        
        ' Verificar se foram encontrados valores correspondentes
        If index > 1 Then
            ' Redimensionar o array para o número correto de valores encontrados
            ReDim Preserve valueArray(1 To index - 1)
            
            ' Classificar os valores em ordem decrescente
            SortArrayDescending valueArray
            
            ' Inserir os valores ordenados na ListBox1
            ListBox1.List = Application.WorksheetFunction.Transpose(valueArray)
        Else
            MsgBox "Nenhum recebimento foi encontrado para a data informada. Insira uma data válida.", vbInformation
        End If
    End If
End Sub

Private Sub CommandButton4_Click()
    Application.Visible = True ' torna o Excel visível novamente
    ThisWorkbook.Activate ' ativa a janela da pasta de trabalho atual
    ThisWorkbook.Windows(1).Visible = True ' torna a janela da pasta de trabalho atual visível
    ThisWorkbook.Windows(1).Activate ' ativa a janela da pasta de trabalho atual
    
    ' ativa a edição de planilhas
    Application.EnableEvents = True
    Application.Interactive = True
    Application.ScreenUpdating = True
    
    'Esconder formulário atual
    Unload Me
End Sub

Private Sub SortArrayDescending(ByRef arr() As Variant)
    ' Implementar uma classificação simples para ordenar o array em ordem decrescente
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Set ws = Nothing
    Set ws2 = Nothing
    Set tbl = Nothing
    Set rngRecebID = Nothing
    Set rngData = Nothing

    Unload Me

End Sub