Private Sub UserForm_Activate()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If

    If TextBox1.value = "" Then
        Unload Me
        CaminhaoLista.Show
    End If
    
    ComboBox1.Clear
    ComboBox2.Clear
    ComboBox3.Clear
    ComboBox4.Clear
    ComboBox5.Clear
    TextBox2.value = ""
    
    ComboBox1.AddItem "Balizamento"
    ComboBox1.AddItem "Drenagem Externa"
    ComboBox1.AddItem "Estucagem"
    ComboBox1.AddItem "Marquise"
    ComboBox1.AddItem "Poço de Visita"
    ComboBox1.AddItem "Reboco"
    ComboBox1.AddItem "Regularização de Calha"
    ComboBox1.AddItem "Regularização de Piso"
    ComboBox1.AddItem "Requadramento de Paredes"
    
    ComboBox2.AddItem "A"
    ComboBox2.AddItem "B"
    ComboBox2.AddItem "Geral"
    
    ComboBox3.AddItem 10
    ComboBox3.AddItem 9
    ComboBox3.AddItem 8
    ComboBox3.AddItem 7
    ComboBox3.AddItem 6
    ComboBox3.AddItem 5
    ComboBox3.AddItem 4
    ComboBox3.AddItem 3
    ComboBox3.AddItem 2
    ComboBox3.AddItem 1
    ComboBox3.AddItem "Equipamentos Comunitários"
    ComboBox3.AddItem "Infraestrutura"
    
    ComboBox4.AddItem "Térreo"
    ComboBox4.AddItem "Primeiro"
    ComboBox4.AddItem "Segundo"
    ComboBox4.AddItem "Terceiro"
    ComboBox4.AddItem "Cobertura"
End Sub

Private Sub ComboBox3_Change()

    If ComboBox3.Value = "Equipamentos Comunitários" Then

        ComboBox5.Clear
        ComboBox5.Value = ""
        ComboBox2.Value = "Geral"
        ComboBox4.Value = "Térreo"
        ComboBox5.AddItem "Quadra"
        ComboBox5.AddItem "Salão de Festas"
        ComboBox5.AddItem "Churrasqueira"
        ComboBox5.AddItem "Piscina"

    ElseIf ComboBox3.Value = "Infraestrutura" Then

        ComboBox5.Clear
        ComboBox5.Value = ""
        ComboBox2.Value = "Geral"
        ComboBox4.Value = "Térreo"
        ComboBox5.AddItem "Muro Limitrofe"
        ComboBox5.AddItem "Sarjeta"
        ComboBox5.AddItem "Meio-fio"
        ComboBox5.AddItem "Pavimentação"

    Else

        ComboBox5.Clear
        ComboBox5.Value = ""
        ComboBox5.AddItem 1
        ComboBox5.AddItem 2
        ComboBox5.AddItem 3
        ComboBox5.AddItem 4
        ComboBox5.AddItem 5
        ComboBox5.AddItem 6
        ComboBox5.AddItem 7
        ComboBox5.AddItem 8
        ComboBox5.AddItem 101
        ComboBox5.AddItem 102
        ComboBox5.AddItem 103
        ComboBox5.AddItem 104
        ComboBox5.AddItem 105
        ComboBox5.AddItem 106
        ComboBox5.AddItem 107
        ComboBox5.AddItem 108
        ComboBox5.AddItem 201
        ComboBox5.AddItem 202
        ComboBox5.AddItem 203
        ComboBox5.AddItem 204
        ComboBox5.AddItem 205
        ComboBox5.AddItem 206
        ComboBox5.AddItem 207
        ComboBox5.AddItem 208
        ComboBox5.AddItem 301
        ComboBox5.AddItem 302
        ComboBox5.AddItem 303
        ComboBox5.AddItem 304
        ComboBox5.AddItem 305
        ComboBox5.AddItem 306
        ComboBox5.AddItem 307
        ComboBox5.AddItem 308
        ComboBox5.AddItem "Cobertura"

    End If

End Sub

Private Sub ComboBox5_Change()
    Dim selectedValue As String
    Dim lastRight As Integer
    Dim lastLeft As Integer
    
    selectedValue = ComboBox5.Value
    lastRight = Val(Right(selectedValue, 1))
    lastLeft = Val(Left(selectedValue, 1))
    
    If Val(lastRight) = 4 Or Val(lastRight) = 3 Or Val(lastRight) = 2 Or Val(lastRight) = 1 Then
        ComboBox2.Value = "A"
    ElseIf Val(lastRight) = 8 Or Val(lastRight) = 7 Or Val(lastRight) = 6 Or Val(lastRight) = 5 Then
        ComboBox2.Value = "B"
    Else
        ComboBox2.value = "Geral"
    End If    

    If lastLeft = 1 And Len(selectedValue) = 3 Then
        ComboBox4.Value = "Primeiro"
    ElseIf lastLeft = 2 And Len(selectedValue) = 3 Then
        ComboBox4.Value = "Segundo"
    ElseIf lastLeft = 3 And Len(selectedValue) = 3 Then
        ComboBox4.Value = "Terceiro"
    ElseIf selectedValue = "Cobertura" Then
        ComboBox4.Value = "Cobertura"
        ComboBox2.Value = "Geral"
    Else
        ComboBox4.Value = "Térreo"
    End If

End Sub


Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxValue As Variant
    Dim validValues As Boolean
    
    ' Definir a planilha "Consumo"
    Set ws = ThisWorkbook.Worksheets("Consumo")
    
    ' Encontrar a última linha preenchida na coluna A da planilha
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar se os valores dos ComboBoxes são válidos
    validValues = ValidateComboBoxValues()
    
    If Not validValues Then
        MsgBox "Preencha todos os campos com valores válidos.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se o valor de TextBox2 é um número válido
    If Not IsNumeric(TextBox2.value) Then
        MsgBox "Preencha um volume válido de consumo.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se há limite no caminhão
    If Val(TextBox2.value) > (Val(TextBox3.value) - Val(TextBox4.value)) Then
        MsgBox "Consumo superior ao limite do caminhão, reveja o consumo.", vbExclamation
        Exit Sub
    End If
    
    ' Obter o maior valor da coluna C e incrementá-lo em 1
    maxValue = Application.WorksheetFunction.Max(ws.Range("C:C")) + 1
    
    ' Inserir os valores dos controles nas colunas desejadas
    ws.Cells(lastRow + 1, "A").value = TextBox1.value
    ws.Cells(lastRow + 1, "C").value = maxValue
    ws.Cells(lastRow + 1, "F").value = ComboBox1.value
    ws.Cells(lastRow + 1, "G").value = ComboBox2.value
    ws.Cells(lastRow + 1, "H").value = ComboBox3.value
    ws.Cells(lastRow + 1, "I").value = ComboBox4.value
    ws.Cells(lastRow + 1, "J").value = ComboBox5.value
    ws.Cells(lastRow + 1, "K").value = TextBox2.value
    ws.Cells(lastRow + 1, "O").value = usernameGlobal
    ws.Cells(lastRow + 1, "P").value = Now()
    
    ' Limpar os controles do formulário
    TextBox1.value = ""
    ComboBox1.value = ""
    ComboBox2.value = ""
    ComboBox3.value = ""
    ComboBox4.value = ""
    ComboBox5.value = ""
    TextBox2.value = ""
    
    ' Exibir mensagem de sucesso
    MsgBox "Valores inseridos com sucesso!", vbInformation
    
    Unload Me
    CaminhaoLista.Show
    
End Sub

Private Function ValidateComboBoxValues() As Boolean
    Dim validValues As Boolean
    
    validValues = True
    
    ' Verificar se os ComboBoxes têm valores válidos
    If Not IsValueInComboBox(ComboBox1, ComboBox1.value) Then validValues = False
    If Not IsValueInComboBox(ComboBox2, ComboBox2.value) Then validValues = False
    If Not IsValueInComboBox(ComboBox3, ComboBox3.value) Then validValues = False
    If Not IsValueInComboBox(ComboBox4, ComboBox4.value) Then validValues = False
    If Not IsValueInComboBox(ComboBox5, ComboBox5.value) Then validValues = False
    
    ValidateComboBoxValues = validValues
End Function

Private Function IsValueInComboBox(comboBox As MSForms.comboBox, value As String) As Boolean
    Dim item As Variant
    
    If Not IsNull(value) Then
        For Each item In comboBox.List
            If Not IsNull(item) Then
                If StrComp(CStr(item), value, vbTextCompare) = 0 Then
                    IsValueInComboBox = True
                    Exit Function
                End If
            End If
        Next item
    End If
    
    IsValueInComboBox = False
End Function



Private Sub CommandButton2_Click()

    Unload Me
    CaminhaoLista.Show

End Sub