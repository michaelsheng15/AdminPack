VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnAddAssessment_Click()
    AssessmentCreate = True
    frmAddAssessment.Show
End Sub

Private Sub btnCompleted_Click()
    ThisWorkbook.Sheets("COMPLETED").Activate
End Sub

Private Sub btnHome_Click()
    ThisWorkbook.Sheets("INTRO").Activate
End Sub


