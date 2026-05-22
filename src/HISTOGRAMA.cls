VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HISTOGRAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Run_Bubble_Sort_Click()
' Execute actions when button at spreadsheet HISTOGRAMA is pressed
    Load SimulationConditionsInputForm
    SimulationConditionsInputForm.Show
End Sub

Private Sub Worksheet_Activate()
' Scroll to first row and first column on worksheet activation
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
End Sub


