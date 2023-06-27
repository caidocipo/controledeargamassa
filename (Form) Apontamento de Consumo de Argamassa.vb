    Private Sub UserForm_Terminate()
    
        ' Oculta o formulário atual
        Unload Me
        ' Carrega e exibe o formulário ListaCaminhoes
        ListaCaminhoes.Show vbModal
        
    End Sub
    
    Private Sub CommandButton2_Click()
    
        ' Oculta o formulário atual
        Unload Me
        ' Carrega e exibe o formulário ListaCaminhoes
        ListaCaminhoes.Show vbModal
        
    End Sub
    
    Private Sub UserForm_Activate()
    
        Dim ws2 As Worksheet
        Dim lastRow2 As Long
        Dim cell2 As Range
        Dim Combo1 As Collection
        Dim Combo2 As Collection
        Dim Combo3 As Collection
        Dim Combo4 As Collection
        Dim Combo5 As Collection
        
        Set ws2 = ThisWorkbook.Sheets("Projetado") ' Substitua "Projetado" pelo nome da sua planilha
        
        ' Inicializa a coleção para armazenar os valores únicos
        Set Combo1 = New Collection
        Set Combo2 = New Collection
        Set Combo3 = New Collection
        Set Combo4 = New Collection
        Set Combo5 = New Collection
        
        ' Obtém a última linha preenchida na coluna A
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        
        'Percorrendo para ComboBox1
        
        ' Percorre os valores da coluna A a partir da linha 2
        For Each cell2 In ws2.Range("A2:A" & lastRow2)
            ' Verifica se o valor já está na coleção antes de adicioná-lo
            On Error Resume Next
            Combo1.Add cell2.value, CStr(cell2.value)
            On Error GoTo 0
        Next cell2
        
        ' Preenche a ComboBox1 com os valores únicos
        For Each Item In Combo1
            ComboBox1.AddItem Item
        Next Item
        
        'Percorrendo para ComboBox2
        
        ' Percorre os valores da coluna B a partir da linha 2
        For Each cell2 In ws2.Range("B2:B" & lastRow2)
            ' Verifica se o valor já está na coleção antes de adicioná-lo
            On Error Resume Next
            Combo2.Add cell2.value, CStr(cell2.value)
            On Error GoTo 0
        Next cell2
        
        ' Preenche a ComboBox1 com os valores únicos
        For Each Item In Combo2
            ComboBox2.AddItem Item
        Next Item
        
        'Percorrendo para ComboBox3
        
        ' Percorre os valores da coluna C a partir da linha 2
        For Each cell2 In ws2.Range("C2:C" & lastRow2)
            ' Verifica se o valor já está na coleção antes de adicioná-lo
            On Error Resume Next
            Combo3.Add cell2.value, CStr(cell2.value)
            On Error GoTo 0
        Next cell2
        
        ' Preenche a ComboBox1 com os valores únicos
        For Each Item In Combo3
            ComboBox3.AddItem Item
        Next Item
        
        'Percorrendo para ComboBox4
        
        ' Percorre os valores da coluna D a partir da linha 2
        For Each cell2 In ws2.Range("D2:D" & lastRow2)
            ' Verifica se o valor já está na coleção antes de adicioná-lo
            On Error Resume Next
            Combo4.Add cell2.value, CStr(cell2.value)
            On Error GoTo 0
        Next cell2
        
        ' Preenche a ComboBox1 com os valores únicos
        For Each Item In Combo4
            ComboBox4.AddItem Item
        Next Item
        
        'Percorrendo para ComboBox5
        
        ' Percorre os valores da coluna E a partir da linha 2
        For Each cell2 In ws2.Range("E2:E" & lastRow2)
            ' Verifica se o valor já está na coleção antes de adicioná-lo
            On Error Resume Next
            Combo5.Add cell2.value, CStr(cell2.value)
            On Error GoTo 0
        Next cell2
        
        ' Preenche a ComboBox1 com os valores únicos
        For Each Item In Combo5
            ComboBox5.AddItem Item
        Next Item
    
    End Sub
    
    Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue1 As Variant
        Dim foundMatch1 As Boolean
        
        ' Obtenha o valor selecionado no ComboBox1
        selectedValue1 = ComboBox1.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox1
        For Each Item In ComboBox1.List
            If Item = selectedValue1 Then
                foundMatch1 = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox1 e exiba uma mensagem de erro
        If Not foundMatch1 Then
            ComboBox1.value = "" ' Limpa o campo ComboBox1
            MsgBox "Selecione um insumo válido.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue2 As Variant
        Dim foundMatch2 As Boolean
        
        ' Obtenha o valor selecionado no ComboBox2
        selectedValue2 = ComboBox2.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox2
        For Each Item In ComboBox2.List
            If Item = selectedValue2 Then
                foundMatch2 = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox2 e exiba uma mensagem de erro
        If Not foundMatch2 Then
            ComboBox2.value = "" ' Limpa o campo ComboBox2
            MsgBox "Selecione um insumo válido.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Private Sub ComboBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue3 As Variant
        Dim foundMatch3 As Boolean
        
        ' Obtenha o valor selecionado no ComboBox3
        selectedValue3 = ComboBox3.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox3
        For Each Item In ComboBox3.List
            If Item = selectedValue3 Then
                foundMatch3 = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox3 e exiba uma mensagem de erro
        If Not foundMatch3 Then
            ComboBox3.value = "" ' Limpa o campo ComboBox3
            MsgBox "Selecione um insumo válido.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Private Sub ComboBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue4 As Variant
        Dim foundMatch4 As Boolean
        
        ' Obtenha o valor selecionado no ComboBox4
        selectedValue4 = ComboBox4.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox4
        For Each Item In ComboBox4.List
            If Item = selectedValue4 Then
                foundMatch4 = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox4 e exiba uma mensagem de erro
        If Not foundMatch4 Then
            ComboBox4.value = "" ' Limpa o campo ComboBox4
            MsgBox "Selecione um insumo válido.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Private Sub ComboBox5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue5 As Variant
        Dim foundMatch5 As Boolean
        
        ' Obtenha o valor selecionado no ComboBox5
        selectedValue5 = ComboBox5.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox5
        For Each Item In ComboBox5.List
            If Item = selectedValue5 Then
                foundMatch5 = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox5 e exiba uma mensagem de erro
        If Not foundMatch5 Then
            ComboBox5.value = "" ' Limpa o campo ComboBox5
            MsgBox "Selecione um insumo válido.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub

    Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o caractere digitado é um número ou uma vírgula
        If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 44 Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub TextBox10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim userInput As String
        
        ' Obtenha o valor digitado no campo TextBox10
        userInput = TextBox10.value
        
        ' Verifique se é um número válido
        If Not IsNumeric(userInput) Then
            MsgBox "O valor digitado não é um número válido. Digite um número válido.", vbExclamation, "Número Inválido"
            TextBox10.value = "" ' Limpa o campo TextBox10
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Private Sub CommandButton1_Click()
        Dim ws As Worksheet
        Dim lastRow As Long
        Dim nextRow As Long
        
        ' Definir a planilha "Consumo"
        Set ws = ThisWorkbook.Worksheets("Consumo")
        
        ' Encontrar a última linha preenchida na coluna RecebID
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Calcular o próximo número de ordem
        nextRow = Application.WorksheetFunction.Max(ws.Range("C2:C" & lastRow)) + 1
        
        ' Inserir os valores na nova linha
        ws.Cells(lastRow + 1, "A").value = TextBox11.value ' RecebID
        ws.Cells(lastRow + 1, "C").value = nextRow ' Ordem
        ws.Cells(lastRow + 1, "F").value = ComboBox1.value ' Serviço
        ws.Cells(lastRow + 1, "G").value = ComboBox3.value ' Coluna
        ws.Cells(lastRow + 1, "H").value = ComboBox2.value ' Bloco
        ws.Cells(lastRow + 1, "I").value = ComboBox4.value ' Pavimento
        ws.Cells(lastRow + 1, "J").value = ComboBox5.value ' Unidade
        ws.Cells(lastRow + 1, "K").value = TextBox10.value ' Consumido (m³)
        
        ' Limpar os campos de entrada
        TextBox11.value = ""
        TextBox10.value = ""
        ComboBox1.value = ""
        ComboBox2.value = ""
        ComboBox3.value = ""
        ComboBox4.value = ""
        ComboBox5.value = ""
        
        ' Mostrar uma mensagem de sucesso
        MsgBox "Registro adicionado com sucesso.", vbInformation, "Sucesso"
        
        Unload Me
        ListaCaminhoes.Show vbModal
    
    End Sub