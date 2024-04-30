Option Explicit

' ____________________________________________________________________________
Sub check_variable_types()

Dim bleCheck As Boolean

bleCheck = True

If VarType(bleCheck) = vbBoolean Then
  Debug.Print "The variable bleCheck is of type Boolean."
End If

End Sub
