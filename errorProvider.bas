Attribute VB_Name = "errorProvider"
Dim vError As Boolean

Public Function ClearErrors()
vError = False
End Function

Public Function SetError(objeto As Object, texto) As Boolean



    vError = True
    objeto.BackColor = vbRed
    objeto.ToolTipText = UCase(texto)


End Function


Public Function ClearError(objeto As Object)
    vError = False
    objeto.BackColor = vbWhite
    objeto.ToolTipText = Empty

End Function



Public Function hasErrors() As Boolean
hasErrors = vError
End Function
