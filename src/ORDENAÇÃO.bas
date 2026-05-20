Attribute VB_Name = "ORDENAÇÃO"
Option Explicit
Public n As Long 'Tamanho do vetor
Public k As Long 'Número de repetições da ordenação
Public CONT As Long 'Contador

Sub METODO_ORDENACAO() 'Rotina que realiza a ordenação de um vetor de n posições, por k vezes
Attribute METODO_ORDENACAO.VB_ProcData.VB_Invoke_Func = "O\n14"
    'Declaração de variáveis
    Dim i As Long 'Índice do vetor
    Dim j As Long 'Índice do vetor
    Dim Ntrocas As Long 'Número de trocas realizado para ordenar
    Dim Xmax As Long 'Número máximo possível de trocas
    Dim A, B As Single 'Variáveis auxiliares
    Dim V() As Single 'Vetor para realizar a ordenação
    
    Application.ScreenUpdating = False
    HISTOGRAMA.Cells(5, 10).Value = n 'Escreve na célula (5,10) da planilha o tamanho do vetor
    Xmax = n * (n - 1) / 2 'Calcula o número máximo possível de trocas
    HISTOGRAMA.Cells(6, 10).Value = Xmax 'Escreve na célula (6,10) o valor de Xmax
    HISTOGRAMA.Cells(7, 10).Value = k 'Escreve na célula (7,10) o valor de k
    
    'Preenchimento da tabela de frequências na planilha
    HISTOGRAMA.Select
    HISTOGRAMA.Range(Cells(3, 2), Cells(65536, 6)).ClearContents
    For CONT = 0 To Xmax
        HISTOGRAMA.Cells(CONT + 3, 2).Value = CONT
        HISTOGRAMA.Cells(CONT + 3, 3).Value = 0
    Next CONT
    
    'Dimensionamento do vetor V
    ReDim V(n) As Single
    
    'Inicialização de variáveis
    i = 0
    j = 0
    
    'Realização das repetições
    For CONT = 1 To k
        
        'Preenchimento do vetor v
        For i = 1 To n
            Randomize 'Inicializa o gerador de números aleatórios
            B = Rnd() 'Geração de um número aleatório e armazenagem em B
            V(i) = B 'Preenchimento da posição i
            Call NUMERO_ALEATORIO(B) 'Chama rotina para montar tabela de frequência dos números aleatórios gerados
        Next i
        
        'Ordenação do vetor V
        Ntrocas = 0
        For i = 1 To n - 1
            For j = i + 1 To n
                If V(i) > V(j) Then
                    A = V(i)
                    V(i) = V(j)
                    V(j) = A
                    Ntrocas = Ntrocas + 1
                End If
            Next j
        Next i
        HISTOGRAMA.Cells(Ntrocas + 3, 3).Value = HISTOGRAMA.Cells(Ntrocas + 3, 3).Value + 1 'Conta o número de trocas e escreve na planilha
    Next CONT
    Call CALCULOS_ESTATISTICOS 'Chama rotina para realizar cálculo de média, variância e desvio padrão
End Sub

Sub CALCULOS_ESTATISTICOS() 'Rotina que realiza cálculo de média, variância e desvio padrão
    Dim LMAX As Long 'Número da última linha da tabela de frequência
    HISTOGRAMA.Activate
    'Pesquisa da última linha preenchida da tabela de frequência
    If IsEmpty(HISTOGRAMA.Cells(65536, 2)) Then
        HISTOGRAMA.Cells(65536, 2).Select
        Selection.End(xlUp).Select
        LMAX = Selection.Row
    Else
        LMAX = 65536
    End If
    
    'Cálculo da frequência relativa e do produto xi*pi para cálculo da média
    For CONT = 3 To LMAX
        HISTOGRAMA.Cells(CONT, 4).Value = HISTOGRAMA.Cells(CONT, 3) / HISTOGRAMA.Cells(7, 10) 'Cálculo da frequência relativa
        HISTOGRAMA.Cells(CONT, 5).Value = HISTOGRAMA.Cells(CONT, 2) * HISTOGRAMA.Cells(CONT, 4) ' Cálculo do produto xi*pi
    Next CONT
    
    HISTOGRAMA.Cells(8, 10).Select
    ActiveCell.FormulaR1C1 = "=SUM(R3C5:R65536C5)" 'Cálculo da média
    'Cálculo dos termos do somatório para cálculo da variância
    For CONT = 3 To LMAX
        HISTOGRAMA.Cells(CONT, 6).Value = ((HISTOGRAMA.Cells(CONT, 2) - HISTOGRAMA.Cells(8, 10)) ^ 2) * (HISTOGRAMA.Cells(CONT, 4))
    Next CONT
    HISTOGRAMA.Cells(9, 10).Select
    ActiveCell.FormulaR1C1 = "=SUM(R3C6:R65536C6)" 'Cálculo da variância
    HISTOGRAMA.Cells(10, 10).Select
    ActiveCell.FormulaR1C1 = "=SQRT(R[-1]C)" 'Cálculo do desvio padrão

End Sub

Sub NUMERO_ALEATORIO(B As Single) 'Rotina para criar um histograma da distribuição dos números aleatórios gerados
    'Utilizou-se uma largura de intervalo de 0.0001 para gerar uma tabela de distribuição de frequência de ocorrência de números em cada faixa
    Dim ALEATORIO As Single
    Dim LINHA As Long
    ALEATORIO = Round(B, 4)
    LINHA = ALEATORIO / 0.0001
    LINHA = LINHA + 5
    HISTOGRAMA.Cells(LINHA, 17).Value = HISTOGRAMA.Cells(LINHA, 17) + 1
End Sub
