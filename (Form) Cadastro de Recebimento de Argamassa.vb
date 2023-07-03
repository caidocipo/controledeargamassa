Private Sub UserForm_Activate()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If

    ComboBox1.AddItem "Argamassa Estabilizada"
    
    ComboBox2.AddItem "20"
    ComboBox2.AddItem "19"
    ComboBox2.AddItem "18"
    ComboBox2.AddItem "17"
    ComboBox2.AddItem "16"
    ComboBox2.AddItem "15"
    ComboBox2.AddItem "14"
    ComboBox2.AddItem "13"
    ComboBox2.AddItem "12"
    ComboBox2.AddItem "11"
    ComboBox2.AddItem "10"
    ComboBox2.AddItem "9"
    ComboBox2.AddItem "8"
    ComboBox2.AddItem "7"
    ComboBox2.AddItem "6"
    ComboBox2.AddItem "5"
    ComboBox2.AddItem "4"
    ComboBox2.AddItem "3"
    ComboBox2.AddItem "2"
    ComboBox2.AddItem "1"

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox1.Text
    validKeys = "0123456789"
    
    ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
    ' Adicionar o caractere "/" após os dois primeiros dígitos e após os próximos dois dígitos
    If Len(currentValue) = 2 Or Len(currentValue) = 5 Then
        TextBox1.Text = currentValue & "/"
    End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox5.Text
    validKeys = "0123456789"
    
    ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
    ' Adicionar o caractere ":" após os dois primeiros dígitos e após os próximos dois dígitos
    If Len(currentValue) = 2 Then
        TextBox5.Text = currentValue & ":"
    End If
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox3.Text
    validKeys = "0123456789"
    
        ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox4.Text
    validKeys = "0123456789"
    
        ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim keyChar As String
    
    ' Obter o caractere digitado
    keyChar = Chr(KeyAscii)
    
    ' Verificar se o caractere é uma barra, hífen ou espaço
    If keyChar = "/" Or keyChar = "-" Or keyChar = " " Then
        ' Ignorar o caractere
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox6.Text = UCase(TextBox6.Text)
End Sub

Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox7.Text
    validKeys = "0123456789"
    
        ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox9.Text
    validKeys = "0123456789,"
    
        ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TextBox10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox10.Text
    validKeys = "0123456789,"
    
        ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newId As Long
    
    ' Verificar se o valor de TextBox1 contém uma data válida
    If Not IsDate(TextBox1.value) Then
        MsgBox "Insira uma data válida..", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se o valor de TextBox6 está no formato correto
    If Not ValidateTextBox6Value(TextBox6.value) Then
        MsgBox "Insira uma placa válida.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se os valores de TextBox9 e TextBox10 são números válidos
    If Not IsNumeric(TextBox9.value) Or Not IsNumeric(TextBox10.value) Then
        MsgBox "Insira valores de consumo e medição válidos.", vbExclamation
        Exit Sub
    End If
    
    ' Obter a última linha preenchida na coluna B da planilha Recebimento
    Set ws = ThisWorkbook.Worksheets("Recebimento")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Obter o novo ID a ser cadastrado (maior valor da coluna B + 1)
    newId = WorksheetFunction.Max(ws.Range("B2:B" & lastRow)) + 1
    
    ' Inserir os valores na planilha Recebimento
    With ws
        .Cells(lastRow + 1, "B").value = newId
        .Cells(lastRow + 1, "C").value = TextBox1.value
        .Cells(lastRow + 1, "D").value = ComboBox1.value
        .Cells(lastRow + 1, "E").value = TextBox3.value
        .Cells(lastRow + 1, "F").value = TextBox4.value
        .Cells(lastRow + 1, "G").value = TextBox5.value
        .Cells(lastRow + 1, "H").value = TextBox6.value
        .Cells(lastRow + 1, "I").value = TextBox7.value
        .Cells(lastRow + 1, "J").value = ComboBox2.value
        .Cells(lastRow + 1, "K").value = TextBox9.value
        .Cells(lastRow + 1, "M").value = TextBox10.value
        .Cells(lastRow + 1, "N").value = TextBox11.value
        .Cells(lastRow + 1, "O").value = usernameGlobal
        .Cells(lastRow + 1, "P").value = Now()
        
    End With
    
    MsgBox "Valores cadastrados com sucesso!", vbInformation
    
    Unload Me
    CaminhaoLista.Show
    
End Sub

Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim value As Variant
    Dim isValid As Boolean
    Dim itemCount As Long
    Dim i As Long
    
    value = ComboBox1.value
    
    ' Verificar se o valor está presente na lista de itens do ComboBox1
    isValid = False
    For i = 0 To ComboBox1.ListCount - 1
        If StrComp(CStr(ComboBox1.List(i)), value, vbTextCompare) = 0 Then
            isValid = True
            Exit For
        End If
    Next i
    
    If Not isValid Then
        ' Valor inválido, limpar o ComboBox1 e exibir uma mensagem
        ComboBox1.value = ""
        MsgBox "Selecione um insumo válido.", vbExclamation
        ComboBox1.SetFocus ' Definir o foco novamente no ComboBox1
    End If
End Sub

Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim value As Variant
    Dim isValid As Boolean
    Dim itemCount As Long
    Dim i As Long
    
    value = ComboBox2.value
    
    ' Verificar se o valor está presente na lista de itens do ComboBox1
    isValid = False
    For i = 0 To ComboBox2.ListCount - 1
        If StrComp(CStr(ComboBox2.List(i)), value, vbTextCompare) = 0 Then
            isValid = True
            Exit For
        End If
    Next i
    
    If Not isValid Then
        ' Valor inválido, limpar o ComboBox1 e exibir uma mensagem
        ComboBox2.value = ""
        MsgBox "Selecione uma medição válida.", vbExclamation
        ComboBox2.SetFocus ' Definir o foco novamente no ComboBox1
    End If
End Sub

Private Function ValidateComboBoxValue(comboBox As MSForms.comboBox, validValues() As Variant) As Boolean
    Dim value As Variant
    Dim item As Variant
    
    value = comboBox.value
    
    For Each item In validValues
        If StrComp(CStr(item), CStr(value), vbTextCompare) = 0 Then
            ValidateComboBoxValue = True
            Exit Function
        End If
    Next item
    
    ValidateComboBoxValue = False
End Function

Private Function ValidateTextBox6Value(ByVal value As String) As Boolean
    Dim prefix As String
    Dim suffix As String
    
    ' Verificar se a string tem pelo menos 7 caracteres
    If Len(value) <> 7 Then
        ValidateTextBox6Value = False
        Exit Function
    End If
    
    ' Extrair os três primeiros caracteres e os quarto, sexto e sétimo caracteres
    prefix = Left(value, 3)
    suffix = Mid(value, 4, 1) & Mid(value, 6, 2)
    
    ' Verificar se os três primeiros caracteres são letras e se o restante são números
    If IsNumeric(suffix) And IsStringAlpha(prefix) Then
        ValidateTextBox6Value = True
    Else
        ValidateTextBox6Value = False
    End If
End Function

Private Function IsStringAlpha(ByVal value As String) As Boolean
    Dim i As Integer
    Dim charCode As Integer
    
    ' Verificar se todos os caracteres da string são letras
    For i = 1 To Len(value)
        charCode = Asc(UCase(Mid(value, i, 1)))
        If charCode < 65 Or charCode > 90 Then
            IsStringAlpha = False
            Exit Function
        End If
    Next i
    
    IsStringAlpha = True
End Function

Private Sub CommandButton2_Click()

    Unload Me
    CaminhaoLista.Show

End Sub