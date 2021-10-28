' Mod de utilizare
' ----------------
' If Verifica_Daca_Exista_Layer = True Then
'       --- linii cod ---
' End If

Public Function Verifica_Daca_Exista_Layer(ByVal strLayerName As String) As Boolean
    Dim objLayer As AcadLayer
    
    If "" = strLayerName Then Exit Function
    
    On Error Resume Next
    Set objLayer = ThisDrawing.Layers(strLayerName)
        
    If objLayer Is Nothing Then
        Verifica_Daca_Exista_Layer = False
    Else
        Verifica_Daca_Exista_Layer = True
    End If
End Function
