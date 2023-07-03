'Sistema de Login

Private Sub UserForm_Initialize()
    TextBox1.SetFocus
    TextBox1.Value = ""
    TextBox2.Value = ""
    Username.usernameGlobal = ""
End Sub

Private Sub UserForm_Activate()
    TextBox1.SetFocus
    TextBox1.Value = ""
    TextBox2.Value = ""
    Username.usernameGlobal = ""
End Sub

Private Sub CommandButton1_Click()
    Dim Username As String
    Dim password As String
    Dim ws As Worksheet
    Dim loginRange As Range
    Dim foundCell As Range
    
    ' Obter os valores digitados nos campos de usuário e senha
    Username = TextBox1.Value
    password = TextBox2.Value
    
    ' Definir a planilha e o intervalo de pesquisa
    Set ws = ThisWorkbook.Sheets("validUser")
    Set loginRange = ws.Range("A:A")
    
    ' Procurar o nome de usuário na coluna "A"
    Set foundCell = loginRange.Find(Username, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Verificar se o nome de usuário foi encontrado e se a senha corresponde
    If Not foundCell Is Nothing Then
        ' Obter a senha correspondente
        Dim passwordCell As Range
        Set passwordCell = ws.Cells(foundCell.Row, foundCell.Column + 1) ' Assumindo que a senha está na próxima coluna
        
        ' Comparar a senha digitada com a senha correspondente
        If password = passwordCell.Value Then
            usernameGlobal = TextBox1.Value
            MsgBox "Login bem-sucedido! Bem-vindo, " & Username & "!", vbInformation, "Login"
            ' Aqui você pode abrir o formulário principal ou executar qualquer ação desejada
            'Esconder formulário atual
            Me.Hide
            'Mostrar formulário
            WelcomeScreen.Show
        Else
            MsgBox "Senha inválida. Tente novamente.", vbExclamation, "Login"
            ' Limpar o campo de senha
            TextBox2.Value = ""
            ' Definir o foco para o campo de senha
            TextBox2.SetFocus
        End If
    Else
        MsgBox "Usuário não encontrado. Tente novamente.", vbExclamation, "Login"
        ' Limpar os campos de usuário e senha
        TextBox1.Value = ""
        TextBox2.Value = ""
        ' Definir o foco para o campo de usuário
        TextBox1.SetFocus
    End If
End Sub

Private Sub CommandButton2_Click()
    ' Salvar o conteúdo antes de fechar o Excel
    ThisWorkbook.Save
    
    ' Fechar o Excel
    Application.Quit
End Sub

Private Sub UserForm_Terminate()
    ' Salvar a pasta de trabalho
    ThisWorkbook.Save
    
    ' Fechar o Excel
    Application.Quit
End Sub