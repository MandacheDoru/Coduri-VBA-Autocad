Public Sub Creaza_Entitate_VBA_Autocad (ByVal NumeEntitate As String)
    Dim objBlock As AcadBlock
    'Dim objLayer As AcadLayer
  
    On Error GoTo Err_Control
    
    Set objBlock = Nothing
    Set objBlock = ThisDrawing.Blocks.Add(NumeEntitate)
      
    'Set objLayer = Nothing
    'Set objLayer = ThisDrawing.Layers.Add(NumeEntitate)  

Finis:
    Exit Sub
    
Err_Control:
    Select Case Err.Number
        Case -2147024809
            Err.Clear
            Resume Finis
        Case Else
            MsgBox Err.Description
            Exit Sub
    End Select
End Sub
