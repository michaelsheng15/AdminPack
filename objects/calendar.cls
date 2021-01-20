VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnAddMeet_Click()
    frmAddMeet.Show
End Sub

Private Sub btnFinishDay_Click()
    frmFinishDay.Show
End Sub

Private Sub btnHome_Click()
ThisWorkbook.Sheets("INTRO").Activate
End Sub
