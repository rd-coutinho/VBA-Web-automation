' VBA Script

Sub AtualizaABCR()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    
    Dim wb As Workbook
    Set wb = Workbooks("Indice ABCR.xlsm")

    ' Capture the current date with year and month in which the script will be executed by the user
    Dim dataCorrente As String
    dataCorrente = Right(Date, 7)
    
    ' Capture how many worksheets are visible
    Dim qtsPlanilhas As Long
    qtsPlanilhas = 0
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        If ws.Visible Then
            qtsPlanilhas = qtsPlanilhas + 1
        End If
    Next ws
    
    '''''''''''''''''''''''
    ' Loop through the date column and when the value is equal to the current date, opens the http adress associated
    Dim dataRange As Range, i As Range
    Set dataRange = Range("A2", Range("A2").End(xlDown))
    
    For Each i In dataRange
        If Right(i.Value, 7) = dataCorrente Then
            i.Select
            Selection.Offset(0, 1).Select
            
            On Error GoTo ErrorHandler
            
            Workbooks.Add Selection.Value       ' Opens the workbook (http)
            
            Dim wb_ABCR As Workbook
            Set wb_ABCR = Workbooks(ActiveWorkbook.Name)
            
            MsgBox "ABCR workbook succesfully downloaded!", vbInformation, "Aviso"
            
            Exit For
            
        End If
    
    Next i
    
    ' Exit sub to not execute the code below the ErrorHandler in case of no errors
    Exit Sub
    
ErrorHandler:
    MsgBox "Error. Perhaps the ABCR index of this month has not been updated yet. Please, wait.", vbInformation, "Aviso"
    

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub