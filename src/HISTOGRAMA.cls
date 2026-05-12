Private Sub BOTAO_EXEC_ORDENAR_Click() 'Rotina do botão presente na planilha
    Load JANELA_DADOS_ENTRADA 'Carrega janela para entrada de parâmetros
    JANELA_DADOS_ENTRADA.Show
End Sub

Private Sub Worksheet_Activate() 'Rotina para exibição correta da tabela de dados
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub