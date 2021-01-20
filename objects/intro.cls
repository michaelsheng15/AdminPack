VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btn_Calendar_Click()
ThisWorkbook.Sheets("CALENDAR").Activate
End Sub

Private Sub btn_Reminders_Click()
ThisWorkbook.Sheets("REMINDERS").Activate
End Sub

Private Sub btn_Assessments_Click()
ThisWorkbook.Sheets("ASSESSMENTS").Activate
End Sub

Private Sub btn_Deliverables_Click()
ThisWorkbook.Sheets("DELIVERABLES").Activate
End Sub


Private Sub btnCalc_Click()
ThisWorkbook.Sheets("CALCULATORS").Activate
End Sub

Private Sub btnCompleted_Click()
ThisWorkbook.Sheets("COMPLETED").Activate
End Sub
