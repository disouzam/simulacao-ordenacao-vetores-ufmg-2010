Private Sub BOTAO_CANCELAR_Click() 'Rotina do botão Cancelar
    Unload JANELA_DADOS_ENTRADA
End Sub

Private Sub BOTAO_INICIAR_Click() 'Rotina do botão Iniciar testes de ordenação
    Application.ScreenUpdating = True 'Ativa atualização de tela
    
    'Verificação do valor digitado para o tamanho do vetor
    If Not IsNumeric(TEXT_TAMANHO) Then
        MsgBox Prompt:="Digite um valor para o tamanho do vetor.", Title:="Erro"
        TEXT_TAMANHO.Value = Empty
        TEXT_TAMANHO.SetFocus
        Exit Sub
    Else
        If TEXT_TAMANHO.Value > 360 Then 'Aceita-se um tamanho máximo de 360 posições para o vetor por questão de espaço
            MsgBox Prompt:="Digite um valor para o tamanho do vetor menor que 360.", Title:="Erro"
            TEXT_TAMANHO.Value = Empty
            TEXT_TAMANHO.SetFocus
            Exit Sub
        Else
            n = TEXT_TAMANHO.Value 'n recebe o número digitado para o tamanho do vetor
        End If
    End If
    
    'Verificação do valor digitado para o número de repetições
    If Not IsNumeric(TEXT_REPETICOES) Then
        MsgBox Prompt:="Digite um valor para o número de repetições desejado.", Title:="Erro"
        TEXT_REPETICOES.Value = Empty
        TEXT_REPETICOES.SetFocus
        Exit Sub
    Else
        k = TEXT_REPETICOES.Value 'k recebeo o valor digitado para o número de repetições
    End If
    Unload JANELA_DADOS_ENTRADA
    Call METODO_ORDENACAO 'Chama rotina para realizar testes de ordenação de vetores
End Sub