Attribute VB_Name = "Module1"
'Michael Sheng
'20888776

Public ReminderCurrentRow As Integer
Public ReminderCreate As Boolean

Public AssessmentCurrentRow As Integer
Public AssessmentCreate As Boolean

Public DeliverableCurrentRow As Integer
Public DeliverableCreate As Boolean

Public finishAssessmentCurrent As Integer

Public finishDeliverCurrent As Integer

'marco thatallows user toe dit entry in reminders sheet
'ReminderCreate is set to false meaning he not not creating a new entry
Sub OpenEditRow()
    ReminderCreate = False
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    ReminderCurrentRow = currentRow
    
    'fills form with prexisting values from the row
    With frm_Reminders
        '.txt_Date = ActiveSheet.Cells(currentRow, "A")
        .cboClass.Text = ActiveSheet.Cells(currentRow, "C")
        .txt_Task = ActiveSheet.Cells(currentRow, "B")
        .txt_Duedate = ActiveSheet.Cells(currentRow, "D")
        .txt_EstTime = ActiveSheet.Cells(currentRow, "E")
        .txt_Questions = ActiveSheet.Cells(currentRow, "F")
    End With
    
    frm_Reminders.Show
End Sub

'delete function used in reminder, assessments, and deliverables sheet
Sub OpenDeleteRow()
    ReminderCreate = False
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    ReminderCurrentRow = currentRow
    
    'deleted cells in range
    Range(Cells(currentRow, "A"), Cells(currentRow, "K")).Delete
End Sub

'macro that open frmFinishAssessment
Sub OpenFinishAssessment()
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    finishAssessmentCurrent = currentRow
    
    'opens form to update entry with final mark
    frmFinishAssessment.Show
    
    'highlights row green to indicate it is completed
    For i = 1 To 8
        Cells(finishAssessmentCurrent, i).Interior.ColorIndex = 4
    Next i
    
    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Dim iCol As Integer
    
    Dim lastRow As Integer
    lastRow = Sheet.Cells(Rows.Count, 13).End(xlUp).Row
    
    Dim outputSheet As Worksheet
    Set outputSheet = Sheets("COMPLETED")
    
    Dim lastOutputRow As Integer
    lastOutputRow = outputSheet.Cells(Rows.Count, 6).End(xlUp).Row 'finds last row on completed sheet in certain column
    
    'logic that copies code over to te completed sheet when button is pressed
    For iCol = 1 To 3 And 7
        outputSheet.Cells(lastOutputRow + 1, iCol + 5) = Sheet.Cells(currentRow, iCol)
    Next iCol
    
    If outputSheet.Cells(lastOutputRow + 1, "H").Value = "MSCI 100" Then
        outputSheet.Cells(lastOutputRow + 1, "H").Interior.ColorIndex = 43
    End If
    If outputSheet.Cells(lastOutputRow + 1, "H").Value = "MATH 115" Then
        outputSheet.Cells(lastOutputRow + 1, "H").Interior.ColorIndex = 3
    End If
    If outputSheet.Cells(lastOutputRow + 1, "H").Value = "MATH 116" Then
        outputSheet.Cells(lastOutputRow + 1, "H").Interior.ColorIndex = 46
    End If
    If outputSheet.Cells(lastOutputRow + 1, "H").Value = "PHYS 115" Then
        outputSheet.Cells(lastOutputRow + 1, "H").Interior.ColorIndex = 33
    End If
    If outputSheet.Cells(lastOutputRow + 1, "H").Value = "CHE 102" Then
        outputSheet.Cells(lastOutputRow + 1, "H").Interior.ColorIndex = 39
    End If
    
    'formats cells on completed sheet
    outputSheet.Range("F4:H100").HorizontalAlignment = xlLeft
    outputSheet.Range("F4:F100").NumberFormat = "mm/dd/yyyy"
    
End Sub

'macro accessed by complete button from reminders sheet
Sub OpenCompleteRow()
    ReminderCreate = False
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    ReminderCurrentRow = currentRow
    
    'highlights row green to indicate completed
    For i = 1 To 8
        Cells(currentRow, i).Interior.ColorIndex = 4
    Next i
    

    Dim Sheet As Worksheet
    Set Sheet = ActiveSheet
    Set Workbook = ActiveWorkbook
    
    Dim iCol As Integer
    Dim lastRow As Integer
    lastRow = Sheet.Cells(Rows.Count, 14).End(xlUp).Row
    
    Dim outputSheet As Worksheet
    Set outputSheet = Sheets("COMPLETED")
    
    Dim lastOutputRow As Integer
    lastOutputRow = outputSheet.Cells(Rows.Count, 1).End(xlUp).Row 'finds last row on completed sheet in certain column
    
    'copies entry over to completed sheet
    For iCol = 1 To 4
        outputSheet.Cells(lastOutputRow + 1, iCol) = Sheet.Cells(currentRow, iCol)
    Next iCol
    
    'logic to highlihgt class cll a certain colour on the completed shit
    If outputSheet.Cells(lastOutputRow + 1, "C").Value = "MSCI 100" Then
        outputSheet.Cells(lastOutputRow + 1, "C").Interior.ColorIndex = 43
    End If
    If outputSheet.Cells(lastOutputRow + 1, "C").Value = "MATH 115" Then
        outputSheet.Cells(lastOutputRow + 1, "C").Interior.ColorIndex = 3
    End If
    If outputSheet.Cells(lastOutputRow + 1, "C").Value = "MATH 116" Then
        outputSheet.Cells(lastOutputRow + 1, "C").Interior.ColorIndex = 46
    End If
    If outputSheet.Cells(lastOutputRow + 1, "C").Value = "PHYS 115" Then
        outputSheet.Cells(lastOutputRow + 1, "C").Interior.ColorIndex = 33
    End If
    If outputSheet.Cells(lastOutputRow + 1, "C").Value = "CHE 102" Then
        outputSheet.Cells(lastOutputRow + 1, "C").Interior.ColorIndex = 39
    End If
    
    'formats cells on completed sheet
    outputSheet.Range("A4:D100").HorizontalAlignment = xlLeft
    outputSheet.Range("A4:A100").NumberFormat = "mm/dd/yyyy"
    outputSheet.Range("D4:D100").NumberFormat = "mm/dd/yyyy"
    

    
End Sub

'macro accessed by complete button on deliverable sheet
Sub OpenCompleteDeliverable()
    ReminderCreate = False
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    finishDeliverCurrent = currentRow
    
    'form to update actual time and final grade
    frmCompleteDeliverable.Show
    
    'highlights row green to indicate completed
    For i = 1 To 8
        Cells(currentRow, i).Interior.ColorIndex = 4
    Next i

    Set Sheet = ActiveSheet
    Dim iCol As Integer

    Dim outputSheet As Worksheet
    Set outputSheet = Sheets("COMPLETED")
    
    Dim lastOutputRow As Integer
    lastOutputRow = outputSheet.Cells(Rows.Count, "L").End(xlUp).Row 'finds last row on completed sheet in certain column
    
    'logic that copies columns to completed sheet
    For iCol = 1 To 3
        outputSheet.Cells(lastOutputRow + 1, iCol + 9) = Sheet.Cells(currentRow, iCol)
    Next iCol

    If outputSheet.Cells(lastOutputRow + 1, "K").Value = "MSCI 100" Then
        outputSheet.Cells(lastOutputRow + 1, "K").Interior.ColorIndex = 43
    End If
    If outputSheet.Cells(lastOutputRow + 1, "K").Value = "MATH 115" Then
        outputSheet.Cells(lastOutputRow + 1, "K").Interior.ColorIndex = 3
    End If
    If outputSheet.Cells(lastOutputRow + 1, "K").Value = "MATH 116" Then
        outputSheet.Cells(lastOutputRow + 1, "K").Interior.ColorIndex = 46
    End If
    If outputSheet.Cells(lastOutputRow + 1, "K").Value = "PHYS 115" Then
        outputSheet.Cells(lastOutputRow + 1, "K").Interior.ColorIndex = 33
    End If
    If outputSheet.Cells(lastOutputRow + 1, "K").Value = "CHE 102" Then
        outputSheet.Cells(lastOutputRow + 1, "K").Interior.ColorIndex = 39
    End If
    
    'formats cells on completed sheet
    outputSheet.Range("J4:L100").HorizontalAlignment = xlLeft
    outputSheet.Range("L4:L100").NumberFormat = "mm/dd/yyyy"
    

    
End Sub


'macro to edit entries on assessment sheet
Sub OpenEditAssessment()
    AssessmentCreate = False 'AssessmentCreate is false to indicate user is not creating new entry
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    AssessmentCurrentRow = currentRow
    
    'fills form with existing values to edit
    With frmAddAssessment
        .txtAssessmentDate = ActiveSheet.Cells(currentRow, "A")
        .txtAssessmentName.Text = ActiveSheet.Cells(currentRow, "B")
        .cboAssessmentClass = ActiveSheet.Cells(currentRow, "C")
        .txtAssessmentLocation = ActiveSheet.Cells(currentRow, "D")
        .txtAssessmentWeight = ActiveSheet.Cells(currentRow, "E")
        .txtGoal = ActiveSheet.Cells(currentRow, "F")
        .txtStudyTime = ActiveSheet.Cells(currentRow, "H")
    End With
    
    'shows form
    frmAddAssessment.Show
End Sub


'macro to edits entries on deliverable sheet
Sub OpenEditDeliverable()
    DeliverableCreate = False 'sets DeliverableCreate to false to indicae user is not creating new entry
    Dim BtnClicked As Object, currentRow As Integer
    Set BtnClicked = ActiveSheet.Buttons(Application.Caller)
    With BtnClicked.TopLeftCell
        currentRow = .Row
    End With
    DeliverableCurrentRow = currentRow
    
    'fills form with existing values
    With frmAddDeliverable
        .txtDeliverableName = ActiveSheet.Cells(currentRow, "A")
        .cboClass = ActiveSheet.Cells(currentRow, "B")
        .txtDeliverableDate = ActiveSheet.Cells(currentRow, "c")
        .txtEst = ActiveSheet.Cells(currentRow, "D")
        '.txtAct = ActiveSheet.Cells(currentRow, "E")
        .txtComments = ActiveSheet.Cells(currentRow, "G")
    End With
    
    'shows form
    frmAddDeliverable.Show
End Sub


