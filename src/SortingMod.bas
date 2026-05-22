Attribute VB_Name = "SortingMod"
Option Explicit
Public n As Long ' Array size
Public k As Long ' Number of samples to be sorted (k arrays of size n)
Public CONT As Long 'Global Counter

Sub METODO_ORDENACAO()
Attribute METODO_ORDENACAO.VB_ProcData.VB_Invoke_Func = "O\n14"
' Sort k arrays of size n
    
    Dim i As Long ' Array index
    Dim j As Long ' Array index
    Dim Ntrocas As Long ' Number of swaps required to sort
    Dim Xmax As Long ' Maximum number of swaps required to sort
    Dim A, B As Single ' Auxiliary variables
    Dim V() As Single ' Array to perform the sorting
    
    Application.ScreenUpdating = False
    HistogramSheet.Cells(5, 10).Value = n ' Write in cell from row 5, column 10 of HistogramSheet the array size chosen
    Xmax = n * (n - 1) / 2 ' Calculation of the maximum number of swaps
    HistogramSheet.Cells(6, 10).Value = Xmax ' Write in cell from row 6, column 10 the value of Xmax
    HistogramSheet.Cells(7, 10).Value = k ' Write in cell from row 7, column 10 the value of k
    
    ' Fill the table with frequencies
    HistogramSheet.Select
    HistogramSheet.Range(Cells(3, 2), Cells(65536, 6)).ClearContents
    For CONT = 0 To Xmax
        HistogramSheet.Cells(CONT + 3, 2).Value = CONT
        HistogramSheet.Cells(CONT + 3, 3).Value = 0
    Next CONT
    
    ' Resize of array V
    ReDim V(n) As Single
    
    ' Variable initialization
    i = 0
    j = 0
    
    ' Perform the k sampling of arrays of size n
    For CONT = 1 To k
        
        ' Fill array v
        For i = 1 To n
            Randomize ' Initialize the random number generator
            B = Rnd() ' Get a random number and store in B
            V(i) = B ' Fill position i
            Call NUMERO_ALEATORIO(B) ' Fill table of random number generated
        Next i
        
        ' Sorting array V
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
        
        ' Count the number of swaps and write in the sheet
        HistogramSheet.Cells(Ntrocas + 3, 3).Value = HistogramSheet.Cells(Ntrocas + 3, 3).Value + 1
    Next CONT
    
    ' Call the routine that calculate descriptive statistics - average, variance and standard deviation
    Call CALCULOS_ESTATISTICOS
End Sub

Sub CALCULOS_ESTATISTICOS()
' Perform descriptive statistical calculations - average, variance and standard deviation

    Dim LMAX As Long ' Last filled row number in frequency table
    HistogramSheet.Activate
    
    ' Retrieve the number of last filled row number
    If IsEmpty(HistogramSheet.Cells(65536, 2)) Then
        HistogramSheet.Cells(65536, 2).Select
        Selection.End(xlUp).Select
        LMAX = Selection.Row
    Else
        LMAX = 65536
    End If
    
    ' Calculation of relative frequency and of the product x_i * p_i for the average value calculation
    For CONT = 3 To LMAX
        ' Relative frequency calculation
        HistogramSheet.Cells(CONT, 4).Va0lue = HistogramSheet.Cells(CONT, 3) / HistogramSheet.Cells(7, 10)
        
        ' Calculation of product x_i * p_i
        HistogramSheet.Cells(CONT, 5).Value = HistogramSheet.Cells(CONT, 2) * HistogramSheet.Cells(CONT, 4)
    Next CONT
    
    HistogramSheet.Cells(8, 10).Select
    
    ' Average calculation
    ActiveCell.FormulaR1C1 = "=SUM(R3C5:R65536C5)"
    
    ' Calculation of terms of summation to calculate variance
    For CONT = 3 To LMAX
        HistogramSheet.Cells(CONT, 6).Value = ((HistogramSheet.Cells(CONT, 2) - HistogramSheet.Cells(8, 10)) ^ 2) * (HistogramSheet.Cells(CONT, 4))
    Next CONT
    HistogramSheet.Cells(9, 10).Select
    
    ' Variance calculation
    ActiveCell.FormulaR1C1 = "=SUM(R3C6:R65536C6)"
    HistogramSheet.Cells(10, 10).Select
    
    ' Standard deviation calculation
    ActiveCell.FormulaR1C1 = "=SQRT(R[-1]C)"

End Sub

Sub NUMERO_ALEATORIO(B As Single)
' Routine to create a histogram of random number generated histogram
    
    ' Used an interval width of 0.0001 to generate frequency distribution table for numbers in each range
    Dim ALEATORIO As Single
    Dim LINHA As Long
    ALEATORIO = Round(B, 4)
    LINHA = ALEATORIO / 0.0001
    LINHA = LINHA + 5
    HistogramSheet.Cells(LINHA, 17).Value = HistogramSheet.Cells(LINHA, 17) + 1
End Sub
