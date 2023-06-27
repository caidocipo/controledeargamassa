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

'Preenche a lista de insumos e de medições

    Private Sub UserForm_Activate()
    
        ' Adiciona o valor "Argamassa" à ComboBox1
        ComboBox1.AddItem "Argamassa"
        
           ' Adiciona o valor "Argamassa" à ComboBox1
        ComboBox2.AddItem "20"
        ComboBox2.AddItem "19"
        ComboBox2.AddItem "18"
        ComboBox2.AddItem "17"
        ComboBox2.AddItem "16"
        ComboBox2.AddItem "15"
        ComboBox2.AddItem "14"
        ComboBox2.AddItem "13"
        ComboBox2.AddItem "12"
        ComboBox2.AddItem "10"
        ComboBox2.AddItem "09"
        ComboBox2.AddItem "08"
        ComboBox2.AddItem "07"
        ComboBox2.AddItem "06"
        ComboBox2.AddItem "05"
        ComboBox2.AddItem "04"
        ComboBox2.AddItem "03"
        ComboBox2.AddItem "02"
        ComboBox2.AddItem "01"
        
    End Sub

'Permitir somente números e vírgula nos campos de volume

    Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o caractere digitado é um número ou uma vírgula
        If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 44 Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o caractere digitado é um número ou uma vírgula
        If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 44 Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        End If
    End Sub

'Permitir somente números nos campos de Remessa, Lacre e Pedido

    Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o caractere digitado é um número ou uma vírgula
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o caractere digitado é um número ou uma vírgula
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        End If
    End Sub
    
    Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o caractere digitado é um número ou uma vírgula
        If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        End If
    End Sub
    
'Somente letras e números e deixar em maiúsculo no campo do caminhão

    Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Obtém o caractere digitado
        Dim inputChar As String
        inputChar = Chr(KeyAscii)
        
        ' Verifica se o caractere é uma letra ou um número
        If Not (IsNumeric(inputChar) Or (inputChar Like "[A-Za-z]")) Then
            ' Cancela a entrada do caractere
            KeyAscii = 0
        Else
            ' Converte o texto para maiúsculas
            TextBox6.value = UCase(TextBox6.value)
        End If
    End Sub

'Preenche a hora juntamente com o usuário

    Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        ' Verifica se o código ASCII do caractere digitado não é um número (exceto o backspace)
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii <> 8 Then ' Código ASCII para o backspace
                KeyAscii = 0 ' Cancela o caractere digitado
            End If
        End If
    End Sub
    
    Private Sub TextBox5_Change()
        Dim value As String
        Dim formattedValue As String
    
        ' Obtém o valor atual da TextBox
        value = TextBox5.value
        
        ' Remove todos os caracteres não numéricos
        value = Application.WorksheetFunction.Substitute(value, ":", "")
        
        ' Verifica se o valor é um número
        If IsNumeric(value) Then
            ' Adiciona automaticamente o caractere ":" ao digitar
            If Len(value) > 2 Then
                value = Left(value, 2) & ":" & Mid(value, 3)
            End If
        End If
        
        ' Atualiza o valor da TextBox
        TextBox5.value = value
    End Sub

'Cadastra novo caminhão conforme informações inseridas

    Private Sub CommandButton1_Click()
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim lastRow As Long
        Dim lastID As String
        Dim sequentialNumber As Long
        
        ' Definir a planilha de destino
        Set ws = ThisWorkbook.Sheets("Recebimento")
        
        ' Definir a tabela
        Set tbl = ws.ListObjects("recebArgamassa")
        
        ' Encontrar a última linha na tabela
        lastRow = tbl.Range.Rows.Count
        
        ' Obter o último valor da coluna "Ordem" e adicionar 1
        Dim maxOrdem As Long
        maxOrdem = Application.WorksheetFunction.Max(tbl.ListColumns("Ordem").DataBodyRange) + 1
        
        ' Preencher os valores nas colunas correspondentes
        With tbl.DataBodyRange.Rows(lastRow)
            .Cells(1, tbl.ListColumns("Ordem").index).value = maxOrdem
            .Cells(1, tbl.ListColumns("Data").index).value = TextBox1.value
            .Cells(1, tbl.ListColumns("Insumo").index).value = ComboBox1.value
            .Cells(1, tbl.ListColumns("Nota de Remessa").index).value = TextBox3.value
            .Cells(1, tbl.ListColumns("Lacre").index).value = TextBox4.value
            .Cells(1, tbl.ListColumns("Saída da Usina").index).value = TextBox5.value
            .Cells(1, tbl.ListColumns("Caminhão").index).value = TextBox6.value
            .Cells(1, tbl.ListColumns("Pedido").index).value = TextBox7.value
            .Cells(1, tbl.ListColumns("Medição").index).value = ComboBox2.value
            .Cells(1, tbl.ListColumns("Recebido (m³)").index).value = TextBox9.value
            .Cells(1, tbl.ListColumns("Medido (m³)").index).value = TextBox10.value
        End With
        
        ' Limpar os campos do formulário
        TextBox1.value = ""
        ComboBox1.value = ""
        TextBox3.value = ""
        TextBox4.value = ""
        TextBox5.value = ""
        TextBox6.value = ""
        TextBox7.value = ""
        ComboBox2.value = ""
        TextBox9.value = ""
        TextBox10.value = ""
        
        ' Atualizar a tabela
        tbl.Resize tbl.Range.Resize(lastRow + 1)
        
        ' Exibir mensagem de sucesso
        MsgBox "Dados inseridos com sucesso!"
        
        Unload Me
        ListaCaminhoes.Show vbModal
    End Sub

'Funcionalidade que faz as verificações dos campos

    Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim userInput As String
        Dim formattedDate As String
        
        ' Obtenha o valor digitado no campo TextBox1
        userInput = TextBox1.value
        
        ' Verifique se é uma data válida no formato DD/MM/YY
        If IsDate(userInput) Then
            ' Converta para o formato DD/MM/YYYY
            formattedDate = Format(DateValue(userInput), "dd/mm/yyyy")
            
            ' Atualize o valor do campo TextBox1
            TextBox1.value = formattedDate
        Else
            MsgBox "A data digitada não é válida. Digite uma data no formato DD/MM/YY.", vbExclamation, "Data Inválida"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue As Variant
        Dim foundMatch As Boolean
        
        ' Obtenha o valor selecionado no ComboBox1
        selectedValue = ComboBox1.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox1
        For Each Item In ComboBox1.List
            If Item = selectedValue Then
                foundMatch = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox1 e exiba uma mensagem de erro
        If Not foundMatch Then
            ComboBox1.value = "" ' Limpa o campo ComboBox1
            MsgBox "Selecione um insumo válido.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub

    Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim selectedValue As Variant
        Dim foundMatch As Boolean
        
        ' Obtenha o valor selecionado no ComboBox2
        selectedValue = ComboBox2.value
        
        ' Verifique se o valor selecionado existe na lista do ComboBox2
        For Each Item In ComboBox2.List
            If Item = selectedValue Then
                foundMatch = True
                Exit For
            End If
        Next Item
        
        ' Se não houver correspondência, limpe o campo ComboBox2 e exiba uma mensagem de erro
        If Not foundMatch Then
            ComboBox2.value = "" ' Limpa o campo ComboBox2
            MsgBox "Selecione uma medição válida.", vbExclamation, "Valor Inválido"
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub

    Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim userInput As String
        
        ' Obtenha o valor digitado no campo TextBox9
        userInput = TextBox9.value
        
        ' Verifique se é um número válido
        If Not IsNumeric(userInput) Then
            MsgBox "O valor digitado não é um número válido. Digite um número válido.", vbExclamation, "Número Inválido"
            TextBox9.value = "" ' Limpa o campo TextBox9
            Cancel = True ' Impede que o foco seja movido para outro controle
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

    Private Sub TextBox5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim userInput As String
        
        ' Obtenha o valor digitado no campo TextBox5
        userInput = TextBox5.value
        
        ' Verifique se é um horário válido no formato HH:MM
        If Not IsTimeValid(userInput) Then
            MsgBox "O horário digitado não é válido. Digite um horário no formato HH:MM.", vbExclamation, "Horário Inválido"
            TextBox5.value = "" ' Limpa o campo TextBox5
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Function IsTimeValid(ByVal timeString As String) As Boolean
        On Error Resume Next
        Dim testTime As Date
        testTime = TimeValue(timeString)
        IsTimeValid = (Err.Number = 0)
        On Error GoTo 0
    End Function
    
    Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        Dim userInput As String
        
        ' Obtenha o valor digitado no campo TextBox6
        userInput = TextBox6.value
        
        ' Verifique se é uma placa válida no formato XXX#### ou XXX#X##
        If Not IsPlateValid(userInput) Then
            MsgBox "A placa digitada não é válida.", vbExclamation, "Placa Inválida"
            TextBox6.value = "" ' Limpa o campo TextBox6
            Cancel = True ' Impede que o foco seja movido para outro controle
        End If
    End Sub
    
    Function IsPlateValid(ByVal plate As String) As Boolean
        Dim pattern As String
        Dim regex As Object
        
        ' Defina o padrão da expressão regular
        pattern = "^[A-Z]{3}\d{4}$|^[A-Z]{3}\d[A-Z]{1}\d{2}$"
        
        ' Inicialize o objeto RegExp
        Set regex = CreateObject("VBScript.RegExp")
        
        ' Configure a expressão regular
        With regex
            .pattern = pattern
            .IgnoreCase = True ' Ignorar diferenciação entre maiúsculas e minúsculas
        End With
        
        ' Verifique se a placa corresponde ao padrão
        IsPlateValid = regex.Test(plate)
    End Function