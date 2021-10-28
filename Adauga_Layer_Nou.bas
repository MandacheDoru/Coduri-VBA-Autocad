Public Sub Adauga_Layer_Nou(ByVal strLayerName As String)
    Dim objLayer As AcadLayer
    
    If "" = strLayerName Then Exit Sub
    
    On Error Resume Next
    
    Set objLayer = ThisDrawing.Layers(strLayerName)
        
    If objLayer Is Nothing Then
        Set objLayer = ThisDrawing.Layers.Add(strLayerName)
        If objLayer Is Nothing Then '
            MsgBox "Nu pot adauga '" & strLayerName & "'"
        Else
            MsgBox "Layer adaugat: '" & objLayer.Name & "'"
        End If
    Else
        MsgBox "Layer-ul exista deja"
    End If
End Sub
