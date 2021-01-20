VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnAssesments_Click()
ThisWorkbook.Sheets("ASSESSMENTS").Activate
End Sub

Private Sub btnDeliverables_Click()
ThisWorkbook.Sheets("DELIVERABLES").Activate
End Sub

Private Sub btnHome_Click()
    ThisWorkbook.Sheets("INTRO").Activate
End Sub

Private Sub btnReminder_Click()
ThisWorkbook.Sheets("REMINDERS").Activate
End Sub

Private Sub sortReminders()
    Set WS = ActiveSheet
    WS.Range("A4:D100").Sort key1:=Range("A4:A100"), _
        order1:=xlAscending, Header:=xlNo
End Sub



Private Sub btnSort_Click()
    Set WS = ActiveSheet
    
    WS.Range("A4:D100").Sort key1:=Range("A4:A100"), _
        order1:=xlAscending, Header:=xlNo
        
    WS.Range("J4:L100").Sort key1:=Range("L4:L100"), _
        order1:=xlAscending, Header:=xlNo
        
    WS.Range("F4:H100").Sort key1:=Range("F4:F100"), _
        order1:=xlAscending, Header:=xlNo
End Sub


