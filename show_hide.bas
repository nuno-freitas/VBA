Attribute VB_Name = "Module3"
Sub show_hide()

    Dim ws As Worksheet
    Dim lista As Variant
    Dim existe As Boolean
    
    lista = Array("sheet_1", "sheet_2", "sheet_3")
    
    'Check if any of these sheets is hidden
    For Each ws In ThisWorkbook.Worksheets
        If Not IsError(Application.Match(ws.Name, lista, 0)) Then
            If ws.Visible <> xlSheetVisible Then
                existe = True
                Exit For
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = False
    
    If existe Then
        'Show all sheets
        For Each ws In ThisWorkbook.Worksheets
            If Not IsError(Application.Match(ws.Name, lista, 0)) Then
                ws.Visible = xlSheetVisible
            End If
        Next ws
        
        Else
            'Hide all sheets (except the active sheet)
            For Each ws In ThisWorkbook.Worksheets
                If Not IsError(Application.Match(ws.Name, lista, 0)) Then
                    If ws.Name <> ActiveSheet.Name Then
                        ws.Visible = xlSheetHidden
                    End If
                End If
            Next ws
    End If
    
    Application.ScreenUpdating = True
    
End Sub

