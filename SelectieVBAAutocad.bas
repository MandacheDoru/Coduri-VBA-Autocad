Sub SelectieVBAAutocad ()
  Dim sset As AcadSelectionSet
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim objBlock As AcadBlockReference
  Dim objEntity As AcadEntity
  Dim varAtts() As AcadAttributeReference
  
  On Error GoTo Err_Control
  
  Set sset = ThisDrawing.SelectionSets.Add("SS1")

  FilterType(0) = 2
  FilterData(0) = "Punct"

  sset.Select acSelectionSetAll, , , FilterType, FilterData
      
  For Each objEntity In sset
    Set objBlock = objEntity
    If objEntity.HasAttributes Then
      varAtts = objEntity.GetAttributes
      For i = LBound(varAtts) To UBound(varAtts)
          If varAtts(i).TagString = "atribut_cautat" Then
            If CInt(varAtts(i).TextString) > nr Then
              nr = CInt(varAtts(i).TextString)
            End If
          End If
      Next i
    obj.Update
    End If
  Next

  ThisDrawing.SelectionSets("SS1").Delete 
    
Finis:
    Exit Sub
    
Err_Control:
    Select Case Err.Number
        Case -2147024809
            Err.Clear
            ThisDrawing.SelectionSets("SS1").Delete
            Resume Finis
        Case Else
            MsgBox Err.Description
            ThisDrawing.SelectionSets("SS1").Delete
            Exit Sub
    End Select    
End sub
