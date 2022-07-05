Attribute VB_Name = "InsertEnEquation"
Sub enEquation()
Attribute enEquation.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.enEquation"
    Application.Keyboard (1033)
    Selection.OMaths.Add Range:=Selection.Range
End Sub
