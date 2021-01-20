VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnClear1_Click()
    Range("B11:F1000").Clear
End Sub

Private Sub btnClear2_Click()
    Range("H11:L1000").Clear
End Sub

Private Sub btnFinalGradeCalc_Click()
    finalGrade.Show
End Sub

Private Sub btnHome_Click()
ThisWorkbook.Sheets("INTRO").Activate
End Sub

Private Sub btnOverallGrade_Click()
    overallGrade.Show
End Sub
