Public Sub CheckForLayerByIteration()
    Dim objLayer As AcadLayer
    Dim strLayerName As String

    strLayername = InputBox("Enter a Layer name to search for: ")
    If "" = strLayername Then Exit Sub    ' exit if no name entered
 
    For Each objLayer In ThisDrawing.Layers    ' iterate layers 
        If 0 = StrComp(objLayer.name, strLayername, vbTextCompare) Then
            MsgBox "Layer '" & strLayername & "' exists"
            Exit Sub                           ' exit after finding layer
        End If
    Next objLayer
    MsgBox "Layer '" & strLayername & "' does not exist"
End Sub
